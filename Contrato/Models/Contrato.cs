namespace Contrato.Models
{
    public class Contrato
    {
        public Empresa Empresa { get; set; }
        public Empregado Empregado { get; set; }
        public Ctps Ctps { get; set; }
        public Endereco Endereco { get; set; }
    }

    public class Empresa
    {
        public string RazaoSocialEmpresa { get; set; }
        public string CnpjEmpresa { get; set; }
        public string EnderecoEmpresa { get; set; }
        public string ComplementoEmpresa { get; set; }
        public string CepEmpresa { get; set; }
    }

    public class Empregado
    {
        public string NomeEmpregado { get; set; }
        public string NacionalidadeEmpregado { get; set; }
        public string EstadoCivilEmpregado { get; set; }
        public string ProfissaoEmpregado { get; set; }
        public string CpfEmpregado { get; set; }
        public string RgEmpregado { get; set; }
        public string EnderecoEmpregado { get; set; }
        public string ComplementoEmpregado { get; set; }
        public string CepEmpregado { get; set; }
    }

    public class Ctps
    {
        public string NumeroCtps { get; set; }
        public string SerieCtps { get; set; }
        public string UfCtps { get; set; }
        public string ValorHoraCtps { get; set; }
    }

    public class Endereco
    {
        public string Localidade { get; set; }
        public string Uf { get; set; }
    }
}
