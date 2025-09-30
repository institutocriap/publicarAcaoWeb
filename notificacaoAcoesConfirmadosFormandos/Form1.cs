using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Net.Mail;
using System.Net;
using System.Text;
using System.Reflection;
using System.Linq;
using NotificacaoAcoesConfirmadasFormandos.Properties;
using NotificacaoAcoesConfirmadasFormandos.Connects;
using System.Net.Http;
using Newtonsoft.Json;
using static NotificacaoAcoesConfirmadasFormandos.Model.GestorEmailConfirmadosModel;
using System.Threading.Tasks;

namespace NotificacaoAcoesConfirmadasFormandos
{
    public partial class Form1 : Form
    {
        public bool teste;
        public string dataTeste;
        public string data;
        public string versao;
        private List<Formando> alunos;
        public static List<MoodleCourses> courses = new List<MoodleCourses>();
        List<Relatorio> listFormandosNotificados = new List<Relatorio>();

        public Form1()
        {
            InitializeComponent();
            Security.remote();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            teste = true;





            Security.remote();

            Version v = Assembly.GetExecutingAssembly().GetName().Version;
            this.Text += " V." + v.Major.ToString() + "." + v.Minor.ToString() + "." + v.Build.ToString();
            versao = @" <br><font size=""-2"">Controlo de versão: " + " V." + v.Major.ToString() + "." + v.Minor.ToString() + "." + v.Build.ToString() + " Assembly built date: " + System.IO.File.GetLastWriteTime(Assembly.GetExecutingAssembly().Location) + " by rc";

            string[] passedInArgs = Environment.GetCommandLineArgs();

            if (passedInArgs.Contains("-a") || passedInArgs.Contains("-A"))
            {
                // Use ConfigureAwait(false) e não espere pelo resultado na UI thread
                Task.Run(() => EnviarNotificacaoPorcentagemFaltasParaCordenadorAsync())
                    .ContinueWith(t =>
                    {
                        if (t.IsFaulted && t.Exception != null)
                        {
                            // Log do erro sem bloquear a UI
                            Task.Run(() => sendEmail(t.Exception.ToString(), Settings.Default.emailenvio, true, "informatica", ""));
                        }
                        Application.Exit();
                    }, TaskScheduler.FromCurrentSynchronizationContext());

                // Não use Cursor.Current aqui, pois bloqueará a thread de UI
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Melhor abordagem: mudar o cursor antes e depois, sem bloquear a UI
            Cursor = Cursors.WaitCursor;

            // Use ConfigureAwait(false) para evitar retornar à thread de UI desnecessariamente
            Task.Run(() => EnviarNotificacaoPorcentagemFaltasParaCordenadorAsync())
                .ContinueWith(t =>
                {
                    // Volte ao cursor normal quando terminar (na thread de UI)
                    Cursor = Cursors.Default;

                    if (t.IsFaulted && t.Exception != null)
                    {
                        // Log do erro sem bloquear a UI
                        Task.Run(() => sendEmail(t.Exception.ToString(), Settings.Default.emailenvio, true, "informatica", ""));
                    }
                }, TaskScheduler.FromCurrentSynchronizationContext());
        }

        // Torne o método completamente assíncrono
        private async Task EnviarNotificacaoPorcentagemFaltasParaCordenadorAsync()
        {
            try
            {
                // Inicializa a conexão em uma task separada
                await Task.Run(() => Connect.HTlocalConnect.ConnInit()).ConfigureAwait(false);

                string queryFormandos = @"
                    SELECT *
                    FROM (SELECT * FROM TBForFormandos) as Formandos, 
                    (SELECT Codigo_Entidade, Codigo_Pais, Rua, Localidade, Codigo_Postal, SubCodigo_Postal, Descricao_Postal FROM TBGerMoradas where Tipo_Entidade = 3) as Moradas, 
                    (SELECT Codigo_Entidade, Telefone1, Telefone2, Email1, Email2 FROM TBGerContactos  where Tipo_Entidade = 3) as Contactos, 
                    (SELECT a.codigo_curso, a.Numero_Accao, f.codigo_formando, f.formando, a.Ref_Accao, c.Descricao, a.Data_Inicio, a.Data_Fim, c.Tipo_Curso, a.Cod_Tecn_Resp FROM TBForInscricoes i 
                        INNER JOIN TBForFormandos f ON i.codigo_formando = f.codigo_formando and i.confirmado = 1 and i.desistente = 0 
                        INNER JOIN TBForAccoes a ON a.versao_rowid = i.rowid_accao INNER JOIN TBForCursos c ON a.codigo_curso = c.codigo_curso WHERE a.Ref_Accao IN 
                        (SELECT DISTINCT Ref_Accao
                            FROM [secretariaVirtual].[dbo].[AcoesConfirmadas] af
                            INNER JOIN TBForAccoes a ON a.versao_rowid = af.versao_rowid
                            INNER JOIN TBForCursos c ON c.Codigo_Curso = a.Codigo_Curso
                            INNER JOIN TbForTiposCurso tc ON tc.Tipo = c.Tipo_Curso
                            WHERE CAST(af.versao_data AS DATE) = CAST(DATEADD(DAY, -1, GETDATE()) AS DATE))
                    ) as Turmas 
                    where Formandos.versao_rowid = Moradas.Codigo_Entidade AND Formandos.versao_rowid = Contactos.Codigo_Entidade AND Turmas.Codigo_Formando = Formandos.Codigo_Formando 
                    order by Ref_Accao";

                // Use um método totalmente assíncrono para processar os dados
                DataTable dataTableFormandos = new DataTable();

                await Task.Run(() =>
                {
                    using (SqlDataAdapter adapterFormandos = new SqlDataAdapter(queryFormandos, Connect.HTlocalConnect.Conn))
                    {
                        adapterFormandos.Fill(dataTableFormandos);
                    }
                }).ConfigureAwait(false);

                alunos = dataTableFormandos.AsEnumerable().Select(row => new Formando
                {
                    CodigoCurso = row.Field<string>("codigo_curso")?.Trim(),
                    CodigoFormando = row.Field<int?>("Codigo_Formando") ?? 0,
                    RefAccao = row.Field<string>("Ref_Accao")?.Trim(),
                    DataInicio = row.IsNull("Data_Inicio") ? (DateTime?)null : row.Field<DateTime>("Data_Inicio"),
                    DataFim = row.IsNull("Data_Fim") ? (DateTime?)null : row.Field<DateTime>("Data_Fim"),
                    TipoCurso = row.Field<string>("Tipo_Curso")?.Trim(),
                    CodTecnResp = row.Field<string>("Cod_Tecn_Resp")?.Trim(),
                    Descricao = row.Field<string>("Descricao")?.Trim(),
                    NumeroAccao = row.Field<int?>("Numero_Accao") ?? 0,
                    NomeAbreviado = row.Field<string>("Nome_Abreviado")?.Trim(),
                    Email = row.Field<string>("Email1")?.Trim(),
                    Sexo = row.Field<string>("Sexo")?.Trim(),
                    Telefone = row.Field<string>("Telefone1")?.Trim(),
                }).ToList();

                await Task.Run(() => Connect.HTlocalConnect.ConnEnd()).ConfigureAwait(false);

                using (var httpClient = new HttpClient())
                {
                    httpClient.Timeout = TimeSpan.FromSeconds(10);

                    // Processar cada aluno usando chamadas assíncronas
                    foreach (var aluno in alunos)
                    {
                        // Nao envia para formacao a medida
                        if (aluno.RefAccao.StartsWith("FM_") || aluno.RefAccao.StartsWith("A"))
                            continue;

                        // Construir a requisição
                        var request = new HttpRequestMessage(HttpMethod.Post, "http://192.168.1.213:8080/api/email/confirmada-acao-aluno");
                        //var request = new HttpRequestMessage(HttpMethod.Post, "http://localhost:5141/api/email/confirmada-acao-aluno");

                        var content = new StringContent(
                            JsonConvert.SerializeObject(new
                            {
                                RefAccao = aluno.RefAccao.ToString(),
                                Tipo = "confirmada_acao_aluno",
                                Role = "formando",
                                CodHt = aluno.CodigoFormando.ToString(),
                                SessaoId = string.Empty
                            }),
                            Encoding.UTF8,
                            "application/json"
                        );

                        request.Content = content;

                        // Chamada assíncrona ao serviço HTTP
                        var response = await httpClient.SendAsync(request).ConfigureAwait(false);

                        if (response.IsSuccessStatusCode)
                        {
                            var responseContent = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                            var retornoAlunosWrapper = JsonConvert.DeserializeObject<EmailResponseWrapper>(responseContent);
                            var retornoAlunos = retornoAlunosWrapper.Templates;

                            foreach (var retornoAluno in retornoAlunos)
                            {
                                Relatorio relatorioFinal = new Relatorio()
                                {
                                    NomeFormador = aluno.NomeAbreviado,
                                    EmailFormador = aluno.Email,
                                    RefAcao = aluno.RefAccao
                                };

                                // Envio de email e registro de log - sem bloquear a UI
                                await Task.Run(() => sendEmail(retornoAluno.TemplateHtmlFinal, "", false, aluno.Email, "", aluno, retornoAluno.CoordenadorEmail)).ConfigureAwait(false);
                                await Task.Run(() => RegistraLog(aluno.CodigoFormando.ToString(), $"Enviado email de aconfirmação da açãopara o formando via Rotina || Ref Ação: {aluno.RefAccao} ", "Ação Confirmada Formando (Aluno)", aluno.RefAccao)).ConfigureAwait(false);

                                listFormandosNotificados.Add(relatorioFinal);
                            }
                        }
                    }
                }

                // Enviar o relatório final de forma assíncrona
                await Task.Run(() => sendEmailRelatorio()).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                // Log do erro sem bloquear a UI
                await Task.Run(() => sendEmail(ex.ToString(), Settings.Default.emailenvio, true, "informatica", "")).ConfigureAwait(false);
            }
        }

        // Métodos de envio de email devem ser atualizados para ser assíncronos também
        private async Task sendEmailRelatorioAsync()
        {
            try
            {
                NetworkCredential basicCredential = new NetworkCredential(Settings.Default.emailenvio, Settings.Default.passwordemail);
                using (SmtpClient client = new SmtpClient())
                {
                    client.Port = 25;
                    client.Host = "mail.criap.com";
                    client.Timeout = 10000;
                    client.DeliveryMethod = SmtpDeliveryMethod.Network;
                    client.Credentials = basicCredential;
                    client.EnableSsl = false;

                    using (MailMessage mm = new MailMessage())
                    {
                        mm.From = new MailAddress("Instituto CRIAP <" + Settings.Default.emailenvio + "> ");

                        if (!teste)
                        {
                            mm.To.Add("informatica@criap.com");
                            mm.To.Add("tecnicopedagogico@criap.com");
                        }
                        else
                        {
                            mm.To.Add("raphaelcastro@criap.com");
                        }

                        string body = "";

                        // Mostra relatório
                        if (listFormandosNotificados != null && listFormandosNotificados.Count > 0)
                        {
                            StringBuilder relatorio = new StringBuilder();
                            relatorio.AppendLine("Relatório de Formandos Notificados para Ações Confirmadas (alunos):<br><br>");

                            string previousRefAcao = null;

                            foreach (var formando in listFormandosNotificados)
                            {
                                if (previousRefAcao != null && formando.RefAcao != previousRefAcao)
                                {
                                    relatorio.AppendLine("<hr>");
                                }

                                relatorio.AppendLine($"<b>Formando (aluno):</b> {formando.NomeFormador}  |  Email: {formando.EmailFormador}  |  Ref Ação: {formando.RefAcao}<br>");

                                previousRefAcao = formando.RefAcao;
                            }
                            body = relatorio.ToString();
                        }
                        else
                        {
                            body = "Não há ações confirmadas no dia de hoje.";
                        }

                        mm.Subject = "Instituto CRIAP || Relatório - Notificação Formando (aluno) - Ação Confirmada" + data;
                        mm.BodyEncoding = UTF8Encoding.UTF8;
                        mm.IsBodyHtml = true;
                        mm.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure;
                        mm.Body = body + "<br> " + versao;

                        // Envio de e-mail de forma assíncrona para não bloquear
                        await Task.Run(() => client.Send(mm)).ConfigureAwait(false);
                    }
                }
            }
            catch (Exception ex)
            {
                // Log do erro
                await Task.Run(() => sendEmail(ex.ToString(), Settings.Default.emailenvio, true, "informatica", "")).ConfigureAwait(false);
            }
        }

        // Manter o método síncrono original para compatibilidade
        private void sendEmailRelatorio()
        {
            try
            {
                // Chama diretamente o método assíncrono e aguarda, mas fora da thread de UI
                Task.Run(() => sendEmailRelatorioAsync()).Wait();
            }
            catch (Exception ex)
            {
                sendEmail(ex.ToString(), Settings.Default.emailenvio, true, "informatica", "");
            }
        }

        private void sendEmail(string body, string tecnica = "", bool error = false, string emailPessoa = "", string temp = "", Formando acao = null, string coordenadorEmail = "")
        {
            try
            {
                NetworkCredential basicCredential = new NetworkCredential(Settings.Default.emailenvio, Settings.Default.passwordemail);
                SmtpClient client = new SmtpClient();
                client.Port = 25;
                client.Host = "mail.criap.com";
                client.Timeout = 10000;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.Credentials = basicCredential;

                MailMessage mm = new MailMessage();
                mm.From = new MailAddress("Instituto CRIAP <" + Settings.Default.emailenvio + "> ");

                if (!error)
                {
                    if (!teste)
                    {
                        mm.To.Add(emailPessoa);
                        mm.CC.Add("informatica@criap.com");
                        mm.CC.Add("tecnicopedagogico@criap.com");
                        if(coordenadorEmail != "")
                        {
                            mm.ReplyToList.Add(new MailAddress(coordenadorEmail));
                        }
                        else
                        {
                            mm.ReplyToList.Add(new MailAddress("tecnicopedagogico@criap.com"));
                        }
                    }
                    else
                    {
                        mm.To.Add("raphaelcastro@criap.com");
                    }
                }
                else
                {
                    if (!teste)
                        mm.To.Add("informatica@criap.com");
                    else
                        mm.To.Add("raphaelcastro@criap.com");
                }

                mm.Subject = (!teste) ? "Ação Confirmada / " : data + " TESTE - Ação Confirmada - Formandos // ";
                if (!error && acao != null)
                    mm.Subject += (acao.NumeroAccao == null) ? "" : acao.Descricao + " - " + acao.NumeroAccao + "ª Edição";
                mm.BodyEncoding = UTF8Encoding.UTF8;
                mm.IsBodyHtml = true;
                mm.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure;
                mm.Body = body + "<br> " + temp + (teste ? versao : "");
                client.Send(mm);
            }
            catch (Exception)
            {
                // Tratar exceção se necessário
            }
        }

        public void RegistraLog(string id, string mensagem, string menu, string refAcao)
        {
            List<objLogSendersFormador> logSenders = new List<objLogSendersFormador>();
            logSenders.Add(new objLogSendersFormador
            {
                idFormador = id,
                mensagem = mensagem,
                menu = menu,
                refAccao = refAcao
            });
            DataBaseLogSave(logSenders);
        }

        public static void DataBaseLogSave(List<objLogSendersFormador> logSenders)
        {
            if (logSenders.Count > 0)
            {
                string subQuery = "INSERT INTO sv_logs (idFormando, refAcao, dataregisto, registo, menu, username) VALUES ";
                for (int i = 0; i < logSenders.Count; i++)
                {
                    if (i < logSenders.Count - 1)
                        subQuery += "('" + logSenders[i].idFormador + "', '" + logSenders[i].refAccao + "', GETDATE(), '" + logSenders[i].mensagem + "', '" + logSenders[i].menu + "', 'system_rotina'), ";
                    else subQuery += "('" + logSenders[i].idFormador + "', '" + logSenders[i].refAccao + "', GETDATE(), '" + logSenders[i].mensagem + "', '" + logSenders[i].menu + "', 'system_rotina') ";
                }

                Connect.SVlocalConnect.ConnInit();
                using (SqlCommand cmd = new SqlCommand(subQuery, Connect.SVlocalConnect.Conn))
                {
                    cmd.ExecuteNonQuery();
                }
                Connect.SVlocalConnect.ConnEnd();
                Connect.closeAll();
            }
        }
    }
}