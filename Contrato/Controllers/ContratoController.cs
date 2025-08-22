using Contrato.Models;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using iText.IO.Font;
using iText.IO.Font.Constants;
using iText.Kernel.Font;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Globalization;
using System.IO;
using System.Net.Http;
using System.Net.Http.Json;
using System.Text;
using System.Threading.Tasks;


using ContratoModel = Contrato.Models.Contrato;
using iTextDocument = iText.Layout.Document;
using iTextPageSize = iText.Kernel.Geom.PageSize;
using PdfDocLayout = iText.Layout.Document;
using PdfParagraph = iText.Layout.Element.Paragraph;
using PdfTextAlignment = iText.Layout.Properties.TextAlignment;
using SysPath = System.IO.Path;
using WordParagraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using WordText = DocumentFormat.OpenXml.Wordprocessing.Text;


namespace ContratoAPI.Controllers

{
    [ApiController]
    [Route("api/[controller]")]
    public class ContratoController : ControllerBase
    {
        private readonly HttpClient _httpClient;

        public ContratoController()
        {
            _httpClient = new HttpClient();
        }

        [HttpPost("criar-docx")]
        public async Task<IActionResult> CriarDocumentoDocx([FromBody] ContratoModel contrato)
        {
            if (contrato == null) return BadRequest("Contrato inválido");

            await PreencherEnderecoViaCep(contrato);

            string templatePath = SysPath.Combine(Directory.GetCurrentDirectory(), "Templates", "ContratoTemplate.docx");
            if (!System.IO.File.Exists(templatePath))
                return NotFound("Template de contrato não encontrado.");

            byte[] fileBytes;
            using (var mem = new MemoryStream())
            {
                using (var fileStream = new FileStream(templatePath, FileMode.Open, FileAccess.Read))
                    fileStream.CopyTo(mem);

                mem.Position = 0;

                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(mem, true))
                {
                    var body = wordDoc.MainDocumentPart.Document.Body;
                    foreach (var paragraph in body.Elements<WordParagraph>())
                    {
                        string paragraphText = paragraph.InnerText;
                        paragraphText = paragraphText
                            .Replace("{{RazaoSocialEmpresa}}", contrato.Empresa?.RazaoSocialEmpresa ?? "")
                            .Replace("{{CnpjEmpresa}}", contrato.Empresa?.CnpjEmpresa ?? "")
                            .Replace("{{EnderecoEmpresa}}", contrato.Empresa?.EnderecoEmpresa ?? "")
                            .Replace("{{ComplementoEmpresa}}", contrato.Empresa?.ComplementoEmpresa ?? "")
                            .Replace("{{CepEmpresa}}", contrato.Empresa?.CepEmpresa ?? "")
                            .Replace("{{NomeEmpregado}}", contrato.Empregado?.NomeEmpregado ?? "")
                            .Replace("{{NacionalidadeEmpregado}}", contrato.Empregado?.NacionalidadeEmpregado ?? "")
                            .Replace("{{EstadoCivilEmpregado}}", contrato.Empregado?.EstadoCivilEmpregado ?? "")
                            .Replace("{{ProfissaoEmpregado}}", contrato.Empregado?.ProfissaoEmpregado ?? "")
                            .Replace("{{CpfEmpregado}}", contrato.Empregado?.CpfEmpregado ?? "")
                            .Replace("{{RgEmpregado}}", contrato.Empregado?.RgEmpregado ?? "")
                            .Replace("{{EnderecoEmpregado}}", contrato.Empregado?.EnderecoEmpregado ?? "")
                            .Replace("{{ComplementoEmpregado}}", contrato.Empregado?.ComplementoEmpregado ?? "")
                            .Replace("{{CepEmpregado}}", contrato.Empregado?.CepEmpregado ?? "")
                            .Replace("{{NumeroCtps}}", contrato.Ctps?.NumeroCtps ?? "")
                            .Replace("{{SerieCtps}}", contrato.Ctps?.SerieCtps ?? "")
                            .Replace("{{UfCtps}}", contrato.Ctps?.UfCtps ?? "")
                            .Replace("{{ValorHoraCtps}}", contrato.Ctps?.ValorHoraCtps ?? "")
                            .Replace("{{Localidade}}", contrato.Endereco?.Localidade ?? "")
                            .Replace("{{Uf}}", contrato.Endereco?.Uf ?? "")
                            .Replace("{{LocalData}}", contrato.Endereco != null
                                ? $"{contrato.Endereco.Localidade} - {contrato.Endereco.Uf}, {DateTime.Now:dd 'de' MMMM 'de' yyyy}"
                                : DateTime.Now.ToString("dd 'de' MMMM 'de' yyyy"));

                        paragraph.RemoveAllChildren<Run>();
                        paragraph.AppendChild(new Run(new WordText(paragraphText)));
                    }

                    wordDoc.MainDocumentPart.Document.Save();
                }

                fileBytes = mem.ToArray();
            }

            return File(fileBytes,
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        "Contrato.docx");
        }


        [HttpPost("criar-pdf")]
        public IActionResult GerarContratoPdf([FromBody] Contrato.Models.Contrato contrato)
        {
            var mem = new MemoryStream();

            var writer = new PdfWriter(mem);
            var pdfDoc = new PdfDocument(writer);
            var document = new iText.Layout.Document(pdfDoc);

            string dataExtenso = DateTime.Now.ToString("dd 'de' MMMM 'de' yyyy", new CultureInfo("pt-BR"));
            string localUfData = $"{contrato.Endereco.Localidade} - {contrato.Endereco.Uf}, {dataExtenso}";

            decimal valorHora = decimal.TryParse(contrato.Ctps?.ValorHoraCtps, NumberStyles.Any, new CultureInfo("pt-BR"), out var v) ? v : 0;
            string valorExtenso = $"{valorHora:C2}".Replace("R$", "R$ ").Replace(",", ",00");

            string contratoTexto = $@"
Por este instrumento particular, de um lado {contrato.Empresa.RazaoSocialEmpresa} estabelecida nesta cidade de 
{contrato.Endereco.Localidade}-{contrato.Endereco.Uf} no endereço {contrato.Empresa.EnderecoEmpresa}, regularmente inscrita no CNPJ (MF) sob o nº {contrato.Empresa.CnpjEmpresa}, por seu 
representante legal, neste ato designada EMPREGADORA e, de outro lado {contrato.Empregado.NomeEmpregado}, 
{contrato.Empregado.NacionalidadeEmpregado}, {contrato.Empregado.EstadoCivilEmpregado}, {contrato.Empregado.ProfissaoEmpregado}, inscrito(a) no CPF sob o nº {contrato.Empregado.CpfEmpregado}, no RG nº 
{contrato.Empregado.RgEmpregado} e portador da CTPS nº {contrato.Ctps.NumeroCtps}, Série {contrato.Ctps.SerieCtps}, residente e domiciliado(a) à 
{contrato.Empregado.EnderecoEmpregado}, daqui em diante denominado(a) EMPREGADO(a).

EMPREGADOR(a) e EMPREGADO(a) em conjunto como PARTES e, individualmente, 
como PARTE.

Têm como justo e acertado o presente Contrato de Trabalho Intermitente que se regerá 
através das cláusulas abaixo:

1. O(a) EMPREGADO(a) é contratado(a) na modalidade de trabalho intermitente, de 
maneira não contínua, nos termos dos Arts. 443 e seu parágrafo 3º, e artigo 452-A e seus 
parágrafos, da CLT.

2. O horário de trabalho será definido em cada convocação enviada pela EMPREGADORA, 
conforme previsto no art. 452-A, § 1º da CLT. O EMPREGADO(a) somente estará obrigado 
a cumprir a jornada estabelecida na convocação que aceitar, observando o limite legal, 
conforme artigo 452-A, § 1º e artigo 58, ambos da CLT.

3. Fica o EMPREGADO admitido para exercer a função de {contrato.Empregado.ProfissaoEmpregado} com remuneração R$ 
{valorHora:F2} ({valorExtenso}) por hora trabalhada, passível de reajuste e atualizações de 
acordo com a legislação.

4. A EMPREGADORA convocará o(a) EMPREGADO(a) por meio da Plataforma de Gestão 
de Trabalho Intermitente TIO Digital, informando expressamente a jornada a ser cumprida, 
com antecedência de pelo menos três dias.

4.1 Recebida a comunicação o(a) EMPREGADO(a) terá um dia útil para comunicar a 
aceitação ou não da proposta, sendo que seu silêncio representará a recusa.

4.2 A recusa à convocação não pode ser considerada, sob hipótese nenhuma, como 
insubordinação por parte do(a) EMPREGADO(a).

5. Aceita a proposta, a parte que, sem justo motivo, descumprir o ajustado, pagará à outra 
parte, no prazo de trinta dias, multa de 50% (cinquenta por cento) da remuneração que 
seria devida, permitida a compensação em igual prazo.

6. Fica ajustado nos termos que dispõe o §1 do artigo 469, da CLT, que o EMPREGADO(a) 
acatará ordem emanada da EMPREGADORA para a prestação de serviços tanto naquela 
localidade de celebração do Contrato Intermitente ou em localidade diversa, tendo em vista 
a necessidade do serviço.

7. O pagamento será realizado em até um dia útil ao final de cada prestação de serviço 
efetuada na convocação, porém, se o prazo de trabalho for maior que um mês, o 
pagamento será todo o quinto dia útil do mês seguinte trabalhado, conforme previsto no §1º 
do art. 459 da CLT.

8. O período de inatividade não será considerado tempo à disposição da EMPREGADORA, 
podendo o(a) EMPREGADO(a) prestar serviços a outros contratantes.

8. O(a) EMPREGADO(a) tem direito a usufruir de um mês de férias a cada doze meses de 
trabalho, não podendo ser convocado para prestar serviços pelo EMPREGADOR.

9. Em caso de dano causado pelo EMPREGADO(a), fica a EMPREGADORA autorizada a 
efetivar o desconto da importância correspondente ao prejuízo, o qual fará com fundamento 
no § único do artigo 462 da Consolidação das Leis do Trabalho, já que essa possibilidade 
fica expressamente prevista em Contrato.

10. O empregado se compromete a respeitar o regulamento interno da empresa, bem como 
seguir os procedimentos de segurança no trabalho do empregador, com a utilização de EPI, 
quando necessário, ciente de que constitui falta grave a inobservância do que ora se 
estabelece, além das previstas no artigo 482 da CLT.

E, por estarem de pleno acordo, assinam ambas as partes este contrato, em duas vias de 
igual teor na presença das testemunhas abaixo relacionadas.

{localUfData}

_________________________                                    
EMPREGADORA                                                         
__________________                                                  
TESTEMUNHA                                                             
_______________________ 
EMPREGADO
_______________________ 
TESTEMUNHA";

            // Título centralizado e em negrito
            var boldFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
            document.Add(new iText.Layout.Element.Paragraph("MINUTA DE CONTRATO PARA O TRABALHO INTERMITENTE")
                .SetFontSize(14)
                .SetFont(boldFont)
                .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                .SetMarginBottom(20));

            // Corpo do contrato
            document.Add(new iText.Layout.Element.Paragraph(contratoTexto)
                .SetTextAlignment(PdfTextAlignment.JUSTIFIED)
                .SetFontSize(11)
                .SetTextAlignment(iText.Layout.Properties.TextAlignment.JUSTIFIED));

            document.Close();

          
            var fileBytes = mem.ToArray();

            return File(fileBytes, "application/pdf", "ContratoIntermitente.pdf");
        }










        private async Task PreencherEnderecoViaCep(ContratoModel contrato)
        {
            if (!string.IsNullOrEmpty(contrato.Empresa?.CepEmpresa))
            {
                var cepLimpo = contrato.Empresa.CepEmpresa.Replace("-", "").Trim();
                try
                {
                    var endereco = await _httpClient.GetFromJsonAsync<Endereco>($"https://viacep.com.br/ws/{cepLimpo}/json/");
                    contrato.Endereco = endereco ?? new Endereco { Localidade = "", Uf = "" };
                }
                catch
                {
                    contrato.Endereco = new Endereco { Localidade = "", Uf = "" };
                }
            }
        }
    }
}
