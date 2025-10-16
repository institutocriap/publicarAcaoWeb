using System.Collections.Generic;

namespace PublicarAcaoWeb.Model
{
    public class AcaoParaPublicacao
    {
        public string CodigoCurso { get; set; }
        public string Descricao { get; set; }
        public string TipoCurso { get; set; }
        public string RefAccao { get; set; }
        public decimal PercAceitacao { get; set; }
        public string Motivo { get; set; }
        public string DataMotivo { get; set; }
        public string StatusTratamento { get; set; }
        public string ProxDataInicio { get; set; }
        public string ProxRefAccao { get; set; }
        public string ProxEstado { get; set; }
        public string ProxPubWeb { get; set; }
    }

    public class PublicacaoResponse
    {
        public bool Sucesso { get; set; }
        public string Mensagem { get; set; }
        public int LinhasAfetadas { get; set; }
        public List<string> AcoesDespublicadas { get; set; }
    }

    public class ErrorResponse
    {
        public int CodigoErro { get; set; }
        public string MensagemErro { get; set; }
    }

    public class RelatorioPublicacao
    {
        public string RefAccao { get; set; }
        public string Descricao { get; set; }
        public string Acao { get; set; }
        public string Motivo { get; set; }
        public decimal PercAceitacao { get; set; }
        public bool Sucesso { get; set; }
        public string Erro { get; set; }
    }

    public class objLogSendersFormador
    {
        public string mensagem { get; set; }
        public string idFormador { get; set; }
        public string menu { get; set; }
        public string refAccao { get; set; }
    }
}