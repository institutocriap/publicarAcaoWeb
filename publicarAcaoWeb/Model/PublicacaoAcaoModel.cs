using System;

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
        public DateTime? DataMotivo { get; set; }
        public string StatusTratamento { get; set; }
        public string ProxDataInicio { get; set; }
        public string ProxRefAccao { get; set; }
        public string ProxEstado { get; set; }
        public int? ProxPubWeb { get; set; }
    }

    public class PublicacaoRequest
    {
        public string RefAccao { get; set; }
        public string Acao { get; set; } // "publicar" ou "despublicar"
    }

    public class PublicacaoResponse
    {
        public bool Success { get; set; }
        public string Message { get; set; }
    }

    public class RelatorioPublicacao
    {
        public string RefAccao { get; set; }
        public string Descricao { get; set; }
        public string Acao { get; set; } // "Publicada", "Não precisou publicar"
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