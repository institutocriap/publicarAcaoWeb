using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NotificacaoAcoesConfirmadasFormandos.Model
{
    public class GestorEmailConfirmadosModel
    {
        public class Formando
        {
            public int VersaoRowId { get; set; }
            public string CodigoCurso { get; set; }
            public string RefAccao { get; set; }
            public string RefAcaoAnterior { get; set; }
            public int NumeroAccao { get; set; }
            public string TipoCurso { get; set; }
            public DateTime? DataInicio { get; set; }
            public DateTime? DataFim { get; set; }
            public string Descricao { get; set; }
            public string CodTecnResp { get; set; }
            public string DescTipoCurso { get; internal set; }
            public int CodigoFormando { get; set; }
            public string NomeAbreviado { get; set; }
            public string Sexo { get; set; }
            public string Telefone { get; set; }
            public string Email { get; set; }
            public string Modulos { get; set; }
            public string SessaoRowId { get; set; }
        }

        public class Coordenador
        {
            public int CodCordenador { get; set; }
            public string NomeCoordenador { get; set; }
            public string EmailCoordenador { get; set; }
            public string UserCoordenador { get; set; }
            public string EmailDepPedago { get; set; }
            public string ContactoDepPedago { get; set; }

        }
        public class objLogSendersFormador
        {
            public string mensagem { get; set; }

            public string idFormador { get; set; }

            public string menu { get; set; }

            public string refAccao { get; set; }
        }

        public class MoodleCourses
        {
            public int Id { get; set; }
            public string shortname { get; set; }
        }

        public class EmailResponseWrapper
        {
            [JsonProperty("templates")]
            public List<EmailResponse> Templates { get; set; }
        }

        public class EmailResponse
        {
            [JsonProperty("template")]
            public string TemplateHtmlFinal { get; set; }

            [JsonProperty("codHt")]
            public string CodigoFormando { get; set; }

            [JsonProperty("coordenadorEmail")]
            public string CoordenadorEmail { get; set; }
        }

        public class Relatorio
        {
            public string NomeFormador { get; set; }
            public string EmailFormador { get; set; }
            public string RefAcao { get; set; }
            public string Arquivos { get; set; }
        }
    }
}
