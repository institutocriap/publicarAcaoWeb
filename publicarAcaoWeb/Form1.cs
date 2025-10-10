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
using PublicarAcaoWeb.Properties;
using PublicarAcaoWeb.Connects;
using System.Net.Http;
using Newtonsoft.Json;
using System.Threading.Tasks;
using PublicarAcaoWeb.Model;

namespace PublicarAcaoWeb
{
    public partial class Form1 : Form
    {
        public bool teste;
        public string dataTeste;
        public string data;
        public string versao;
        private List<AcaoParaPublicacao> acoesParaProcessar;
        List<RelatorioPublicacao> listAcoesPublicadas = new List<RelatorioPublicacao>();

        // URLs das APIs
        private const string API_PUBLICAR = "http://localhost:5141/api/publicaAcao/publicar";
        private const string API_DESPUBLICAR = "http://localhost:5141/api/publicaAcao/despublicar";

        public Form1()
        {
            InitializeComponent();
            Security.remote();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            teste = true;
            data = DateTime.Now.ToString("dd/MM/yyyy");

            Security.remote();

            Version v = Assembly.GetExecutingAssembly().GetName().Version;
            this.Text += " V." + v.Major.ToString() + "." + v.Minor.ToString() + "." + v.Build.ToString();
            versao = @" <br><font size=""-2"">Controlo de versão: " + " V." + v.Major.ToString() + "." + v.Minor.ToString() + "." + v.Build.ToString() + " Assembly built date: " + System.IO.File.GetLastWriteTime(Assembly.GetExecutingAssembly().Location) + " by rc";

            string[] passedInArgs = Environment.GetCommandLineArgs();

            if (passedInArgs.Contains("-a") || passedInArgs.Contains("-A"))
            {
                // Executa a rotina de publicação de ações automaticamente
                Task.Run(() => ExecutarRotinaPublicacaoAcoesAsync())
                    .ContinueWith(t =>
                    {
                        if (t.IsFaulted && t.Exception != null)
                        {
                            // Log do erro sem bloquear a UI
                            Task.Run(() => sendEmail(t.Exception.ToString(), Settings.Default.emailenvio, true, "informatica", ""));
                        }
                        Application.Exit();
                    }, TaskScheduler.FromCurrentSynchronizationContext());
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Execução manual da rotina
            Cursor = Cursors.WaitCursor;

            Task.Run(() => ExecutarRotinaPublicacaoAcoesAsync())
                .ContinueWith(t =>
                {
                    Cursor = Cursors.Default;

                    if (t.IsFaulted && t.Exception != null)
                    {
                        // Log do erro sem bloquear a UI
                        Task.Run(() => sendEmail(t.Exception.ToString(), Settings.Default.emailenvio, true, "informatica", ""));
                    }
                }, TaskScheduler.FromCurrentSynchronizationContext());
        }

        private async Task ExecutarRotinaPublicacaoAcoesAsync()
        {
            try
            {
                // Inicializa a conexão
                await Task.Run(() => Connect.HTlocalConnect.ConnInit()).ConfigureAwait(false);

                // Busca as ações que precisam ser processadas
                string queryAcoes = @"
                    DECLARE @data_base DATE = CAST(GETDATE() AS DATE); -- Data de hoje

                    -- ======================================================================
                    -- Blocos de estados auditados
                    -- ======================================================================
                    WITH adiados AS (
	                    SELECT versao_rowid_Accao, versao_data
	                    FROM (
		                    SELECT *,
			                       ROW_NUMBER() OVER (PARTITION BY versao_rowid_Accao ORDER BY versao_data DESC) AS rn
		                    FROM TBAuditAccoes
		                    WHERE CAST(versao_data AS DATE) = @data_base
	                    ) o
	                    WHERE rn = 1 AND Codigo_Estado = 2 AND Codigo_Estado_Antigo != 2
                    ),
                    confirmados AS (
	                    SELECT versao_rowid_Accao
	                    FROM (
		                    SELECT *,
			                       ROW_NUMBER() OVER (PARTITION BY versao_rowid_Accao ORDER BY versao_data DESC) AS rn
		                    FROM TBAuditAccoes
		                    WHERE CAST(versao_data AS DATE) = @data_base
	                    ) o
	                    WHERE rn = 1 
	                      AND Codigo_Estado IN (1,5)
	                      AND ((Codigo_Estado = 1 AND Codigo_Estado_Antigo <> 1)
		                    OR (Codigo_Estado = 5 AND Codigo_Estado_Antigo <> 5))
                    ),
                    confirmados_provisorios AS (
	                    SELECT versao_rowid_Accao
	                    FROM (
		                    SELECT *,
			                       ROW_NUMBER() OVER (PARTITION BY versao_rowid_Accao ORDER BY versao_data DESC) AS rn
		                    FROM TBAuditAccoes
		                    WHERE CAST(versao_data AS DATE) = @data_base
	                    ) o
	                    WHERE rn = 1 
	                      AND Codigo_Estado = 13
	                      AND Codigo_Estado_Antigo <> 13
                    ),
                    cancelados AS (
	                    SELECT versao_rowid_Accao
	                    FROM (
		                    SELECT *,
			                       ROW_NUMBER() OVER (PARTITION BY versao_rowid_Accao ORDER BY versao_data DESC) AS rn
		                    FROM TBAuditAccoes
		                    WHERE CAST(versao_data AS DATE) = @data_base
	                    ) o
	                    WHERE rn = 1 AND Codigo_Estado = 3 AND Codigo_Estado_Antigo != 3
                    ),

                    -- ======================================================================
                    -- % de aceitação (sessões validadas / total)
                    -- ======================================================================
                    SessoesValidadas AS (
	                    SELECT s.Rowid_Accao, COUNT(*) AS sessoes_validadas
	                    FROM secretariaVirtual.dbo.agenda_formador_disponibilidade d
	                    JOIN TBForSessoes s ON d.rowid_sessao = s.versao_rowid
	                    WHERE d.historico = 0 AND d.estado IN ('D','A')
	                    GROUP BY s.Rowid_Accao
                    ),
                    TotalSessoes AS (
	                    SELECT Rowid_Accao, COUNT(*) AS numero_de_sessoes
	                    FROM TBForSessoes
	                    WHERE Codigo_Formador NOT IN (15683,15684,1053,1058,1425,704,699)
	                    GROUP BY Rowid_Accao
                    ),
                    Aceitacao AS (
	                    SELECT 
		                    t.Rowid_Accao,
		                    CAST(ROUND(
			                    (CAST(v.sessoes_validadas AS FLOAT) / NULLIF(t.numero_de_sessoes, 0)) * 100, 2
		                    ) AS DECIMAL(5,2)) AS perc_sessoes_validadas
	                    FROM TotalSessoes t
	                    LEFT JOIN SessoesValidadas v ON t.Rowid_Accao = v.Rowid_Accao
                    ),

                    -- ======================================================================
                    -- Próxima ação futura do mesmo curso
                    -- ======================================================================
                    ProximaAccao AS (
	                    SELECT 
		                    a.Codigo_Curso,
		                    a.Ref_Accao,
		                    (
			                    SELECT TOP 1 a2.Ref_Accao
			                    FROM humantrain.dbo.TBForAccoes a2
			                    JOIN humantrain.dbo.TBForEstados e2 ON e2.Codigo_Estado = a2.Codigo_Estado
			                    LEFT JOIN humantrain.dbo.TBForCandAccoes n2 
				                    ON n2.Codigo_Curso = a2.Codigo_Curso 
			                       AND n2.Numero_Accao = a2.Numero_Accao
			                    WHERE a2.Codigo_Curso = a.Codigo_Curso 
			                      AND a2.Ref_Accao <> a.Ref_Accao
			                      AND a2.Data_Inicio > @data_base
			                      AND e2.Codigo_Estado NOT IN (2,3)
			                    ORDER BY a2.Data_Inicio ASC
		                    ) AS Prox_Ref_Accao,
		                    (
			                    SELECT TOP 1 a2.Data_Inicio
			                    FROM humantrain.dbo.TBForAccoes a2
			                    JOIN humantrain.dbo.TBForEstados e2 ON e2.Codigo_Estado = a2.Codigo_Estado
			                    WHERE a2.Codigo_Curso = a.Codigo_Curso 
			                      AND a2.Ref_Accao <> a.Ref_Accao
			                      AND a2.Data_Inicio > @data_base
			                      AND e2.Codigo_Estado NOT IN (2,3)
			                    ORDER BY a2.Data_Inicio ASC
		                    ) AS Prox_Data_Inicio,
		                    (
			                    SELECT TOP 1 e2.Estado
			                    FROM humantrain.dbo.TBForAccoes a2
			                    JOIN humantrain.dbo.TBForEstados e2 ON e2.Codigo_Estado = a2.Codigo_Estado
			                    WHERE a2.Codigo_Curso = a.Codigo_Curso 
			                      AND a2.Ref_Accao <> a.Ref_Accao
			                      AND a2.Data_Inicio > @data_base
			                      AND e2.Codigo_Estado NOT IN (2,3)
			                    ORDER BY a2.Data_Inicio ASC
		                    ) AS Prox_Estado,
		                    (
			                    SELECT TOP 1 n2.Pub_Web
			                    FROM humantrain.dbo.TBForAccoes a2
			                    JOIN humantrain.dbo.TBForEstados e2 ON e2.Codigo_Estado = a2.Codigo_Estado
			                    LEFT JOIN humantrain.dbo.TBForCandAccoes n2 
				                    ON n2.Codigo_Curso = a2.Codigo_Curso 
			                       AND n2.Numero_Accao = a2.Numero_Accao
			                    WHERE a2.Codigo_Curso = a.Codigo_Curso 
			                      AND a2.Ref_Accao <> a.Ref_Accao
			                      AND a2.Data_Inicio > @data_base
			                      AND e2.Codigo_Estado NOT IN (2,3)
			                    ORDER BY a2.Data_Inicio ASC
		                    ) AS Prox_Pub_Web
	                    FROM humantrain.dbo.TBForAccoes a
                    )

                    -- ======================================================================
                    -- Consulta final consolidada
                    -- ======================================================================
                    SELECT DISTINCT 
	                    c.Codigo_Curso,
	                    c.Descricao,
	                    c.Tipo_Curso,
	                    a.Ref_Accao,
	                    ISNULL(ace.perc_sessoes_validadas, 0) AS Perc_Aceitacao,
	                    CASE 
		                    WHEN adiados.versao_rowid_Accao IS NOT NULL THEN 'Adiado'
		                    WHEN confirmados.versao_rowid_Accao IS NOT NULL THEN 'Confirmado'
		                    WHEN confirmados_provisorios.versao_rowid_Accao IS NOT NULL THEN 'Confirmado Provisório'
		                    WHEN cancelados.versao_rowid_Accao IS NOT NULL THEN 'Cancelado'
		                    WHEN (c.Tipo_Curso IN ('11','12') AND CAST(COALESCE(Qfasecandidatura.qfasedata, TfaseCandidatura.tfasedata, SfaseCandidatura.sfasedata, Pfasecandidatura.pfasedata) AS DATE) = @data_base)
			                      OR (c.Tipo_Curso NOT IN ('11','12') AND CAST(COALESCE(Qfasecandidatura.qfasedata, TfaseCandidatura.tfasedata, SfaseCandidatura.sfasedata, Pfasecandidatura.pfasedata) AS DATE) = @data_base)
		                    THEN 'Fases Caducadas'
		                    ELSE 'Outros'
	                    END AS Motivo,

	                    -- Nova coluna para data que gerou o motivo
	                    CASE
		                    WHEN adiados.versao_rowid_Accao IS NOT NULL THEN adiados.versao_data 
		                    WHEN (c.Tipo_Curso IN ('11','12') AND COALESCE(Qfasecandidatura.qfasedata, TfaseCandidatura.tfasedata, SfaseCandidatura.sfasedata, Pfasecandidatura.pfasedata) = @data_base)
			                      OR (c.Tipo_Curso NOT IN ('11','12') AND COALESCE(Qfasecandidatura.qfasedata, TfaseCandidatura.tfasedata, SfaseCandidatura.sfasedata, Pfasecandidatura.pfasedata) = @data_base)
		                    THEN COALESCE(Qfasecandidatura.qfasedata, TfaseCandidatura.tfasedata, SfaseCandidatura.sfasedata, Pfasecandidatura.pfasedata) 
		                    ELSE NULL
	                    END AS Data_Motivo,

	                    -- Log separado em colunas
	                    CASE WHEN pa.Prox_Ref_Accao IS NULL THEN 'Não Tratado'
		                     WHEN pa.Prox_Pub_Web = 1 THEN 'Já Publicado'
		                     ELSE 'Já Tratado'
	                    END AS Status_Tratamento,
	                    FORMAT(pa.Prox_Data_Inicio, 'dd/MM/yyyy') AS Prox_Data_Inicio,
	                    pa.Prox_Ref_Accao,
	                    pa.Prox_Estado,
	                    pa.Prox_Pub_Web

                    FROM humantrain.dbo.TBForAccoes a
                    JOIN humantrain.dbo.TBForCursos c ON a.Codigo_Curso = c.Codigo_Curso
                    JOIN humantrain.dbo.TBForEstados e ON e.Codigo_Estado = a.Codigo_Estado
                    LEFT JOIN (
	                    SELECT rowid_ecran, valor AS pfasedata FROM humantrain.dbo.TBGerCUValores WHERE nome_campo = 'Pfasecandidatura'
                    ) Pfasecandidatura ON a.versao_rowid = Pfasecandidatura.rowid_ecran
                    LEFT JOIN (
	                    SELECT rowid_ecran, valor AS sfasedata FROM humantrain.dbo.TBGerCUValores WHERE nome_campo = 'Sfasecandidatura'
                    ) SfaseCandidatura ON a.versao_rowid = SfaseCandidatura.rowid_ecran
                    LEFT JOIN (
	                    SELECT rowid_ecran, valor AS tfasedata FROM humantrain.dbo.TBGerCUValores WHERE nome_campo = 'Tfasecandidatura'
                    ) TfaseCandidatura ON a.versao_rowid = TfaseCandidatura.rowid_ecran
                    LEFT JOIN (
	                    SELECT rowid_ecran, valor AS qfasedata FROM humantrain.dbo.TBGerCUValores WHERE nome_campo = 'Ufasecandidatura'
                    ) Qfasecandidatura ON a.versao_rowid = Qfasecandidatura.rowid_ecran
                    LEFT JOIN adiados ON a.versao_rowid = adiados.versao_rowid_Accao
                    LEFT JOIN confirmados ON a.versao_rowid = confirmados.versao_rowid_Accao
                    LEFT JOIN confirmados_provisorios ON a.versao_rowid = confirmados_provisorios.versao_rowid_Accao
                    LEFT JOIN cancelados ON a.versao_rowid = cancelados.versao_rowid_Accao
                    LEFT JOIN Aceitacao ace ON a.versao_rowid = ace.Rowid_Accao
                    LEFT JOIN ProximaAccao pa ON pa.Codigo_Curso = a.Codigo_Curso AND pa.Ref_Accao = a.Ref_Accao
                    WHERE 
                    (
	                    (c.Tipo_Curso IN ('11','12') AND CAST(COALESCE(Qfasecandidatura.qfasedata, TfaseCandidatura.tfasedata, SfaseCandidatura.sfasedata, Pfasecandidatura.pfasedata) AS DATE) = @data_base)
                     OR
	                    (c.Tipo_Curso NOT IN ('11','12') AND CAST(COALESCE(Qfasecandidatura.qfasedata, TfaseCandidatura.tfasedata, SfaseCandidatura.sfasedata, Pfasecandidatura.pfasedata) AS DATE) = @data_base)
                     OR
	                    (adiados.versao_rowid_Accao IS NOT NULL)
                     OR
	                    (confirmados.versao_rowid_Accao IS NOT NULL)
                     OR
	                    (confirmados_provisorios.versao_rowid_Accao IS NOT NULL)
                     OR
	                    (cancelados.versao_rowid_Accao IS NOT NULL)
                    )
                    AND a.Ref_Accao NOT LIKE 'FM_%'
                    ORDER BY c.Descricao;";

                DataTable dataTableAcoes = new DataTable();

                await Task.Run(() =>
                {
                    using (SqlDataAdapter adapterAcoes = new SqlDataAdapter(queryAcoes, Connect.HTlocalConnect.Conn))
                    {
                        adapterAcoes.Fill(dataTableAcoes);
                    }
                }).ConfigureAwait(false);

                // Converte os dados para objetos
                acoesParaProcessar = dataTableAcoes.AsEnumerable().Select(row => new AcaoParaPublicacao
                {
                    CodigoCurso = row.Field<string>("Codigo_Curso")?.Trim(),
                    Descricao = row.Field<string>("Descricao")?.Trim(),
                    TipoCurso = row.Field<string>("Tipo_Curso")?.Trim(),
                    RefAccao = row.Field<string>("Ref_Accao")?.Trim(),
                    PercAceitacao = row.Field<decimal?>("Perc_Aceitacao") ?? 0,
                    Motivo = row.Field<string>("Motivo")?.Trim(),
                    DataMotivo = row.IsNull("Data_Motivo") ? (DateTime?)null : row.Field<DateTime>("Data_Motivo"),
                    StatusTratamento = row.Field<string>("Status_Tratamento")?.Trim(),
                    ProxDataInicio = row.Field<string>("Prox_Data_Inicio")?.Trim(),
                    ProxRefAccao = row.Field<string>("Prox_Ref_Accao")?.Trim(),
                    ProxEstado = row.Field<string>("Prox_Estado")?.Trim(),
                    ProxPubWeb = row.Field<int?>("Prox_Pub_Web")
                }).ToList();

                await Task.Run(() => Connect.HTlocalConnect.ConnEnd()).ConfigureAwait(false);

                // Processa cada ação
                using (var httpClient = new HttpClient())
                {
                    httpClient.Timeout = TimeSpan.FromSeconds(30);

                    foreach (var acao in acoesParaProcessar)
                    {
                        // Processa apenas ações com Status_Tratamento = "Já Tratado"
                        if (acao.StatusTratamento != "Já Tratado")
                        {
                            // Adiciona ao relatório como "não precisou"
                            listAcoesPublicadas.Add(new RelatorioPublicacao
                            {
                                RefAccao = acao.RefAccao,
                                Descricao = acao.Descricao,
                                Acao = "Não processada",
                                Motivo = $"Status: {acao.StatusTratamento}",
                                PercAceitacao = acao.PercAceitacao,
                                Sucesso = true,
                                Erro = ""
                            });
                            continue;
                        }

                        // Para ações "Já Tratado" com mais de 50% de aceitação - não faz nada
                        if (acao.PercAceitacao >= 50)
                        {
                            listAcoesPublicadas.Add(new RelatorioPublicacao
                            {
                                RefAccao = acao.RefAccao,
                                Descricao = acao.Descricao,
                                Acao = "Não precisou publicar",
                                Motivo = $"Aceitação >= 50% ({acao.PercAceitacao}%)",
                                PercAceitacao = acao.PercAceitacao,
                                Sucesso = true,
                                Erro = ""
                            });

                            // Registra log
                            await Task.Run(() => RegistraLog(acao.RefAccao, $"Ação não publicada - Aceitação >= 50% ({acao.PercAceitacao}%) || Ref Ação: {acao.RefAccao}", "Publicação Ação", acao.RefAccao)).ConfigureAwait(false);
                            continue;
                        }

                        // Para ações com menos de 50% de aceitação - publicar a próxima ação
                        if (string.IsNullOrEmpty(acao.ProxRefAccao))
                        {
                            listAcoesPublicadas.Add(new RelatorioPublicacao
                            {
                                RefAccao = acao.RefAccao,
                                Descricao = acao.Descricao,
                                Acao = "Erro - Não há próxima ação",
                                Motivo = $"Aceitação < 50% ({acao.PercAceitacao}%) mas não há próxima ação",
                                PercAceitacao = acao.PercAceitacao,
                                Sucesso = false,
                                Erro = "Próxima ação não encontrada"
                            });
                            continue;
                        }

                        // Chama a API de publicação
                        try
                        {
                            var publicacaoRequest = new PublicacaoRequest
                            {
                                RefAccao = acao.ProxRefAccao,
                                Acao = "publicar"
                            };

                            var request = new HttpRequestMessage(HttpMethod.Post, API_PUBLICAR);
                            var content = new StringContent(
                                JsonConvert.SerializeObject(publicacaoRequest),
                                Encoding.UTF8,
                                "application/json"
                            );
                            request.Content = content;

                            var response = await httpClient.SendAsync(request).ConfigureAwait(false);
                            var responseContent = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

                            if (response.IsSuccessStatusCode)
                            {
                                var publicacaoResponse = JsonConvert.DeserializeObject<PublicacaoResponse>(responseContent);

                                listAcoesPublicadas.Add(new RelatorioPublicacao
                                {
                                    RefAccao = acao.RefAccao,
                                    Descricao = acao.Descricao,
                                    Acao = $"Publicada próxima ação: {acao.ProxRefAccao}",
                                    Motivo = $"Aceitação < 50% ({acao.PercAceitacao}%)",
                                    PercAceitacao = acao.PercAceitacao,
                                    Sucesso = publicacaoResponse?.Success ?? true,
                                    Erro = publicacaoResponse?.Success == false ? publicacaoResponse.Message : ""
                                });

                                // Registra log
                                await Task.Run(() => RegistraLog(acao.RefAccao, $"Publicada próxima ação {acao.ProxRefAccao} - Aceitação < 50% ({acao.PercAceitacao}%) || Ref Ação: {acao.RefAccao}", "Publicação Ação", acao.RefAccao)).ConfigureAwait(false);
                            }
                            else
                            {
                                listAcoesPublicadas.Add(new RelatorioPublicacao
                                {
                                    RefAccao = acao.RefAccao,
                                    Descricao = acao.Descricao,
                                    Acao = "Erro na publicação",
                                    Motivo = $"Aceitação < 50% ({acao.PercAceitacao}%)",
                                    PercAceitacao = acao.PercAceitacao,
                                    Sucesso = false,
                                    Erro = $"HTTP {response.StatusCode}: {responseContent}"
                                });
                            }
                        }
                        catch (Exception ex)
                        {
                            listAcoesPublicadas.Add(new RelatorioPublicacao
                            {
                                RefAccao = acao.RefAccao,
                                Descricao = acao.Descricao,
                                Acao = "Erro na publicação",
                                Motivo = $"Aceitação < 50% ({acao.PercAceitacao}%)",
                                PercAceitacao = acao.PercAceitacao,
                                Sucesso = false,
                                Erro = ex.Message
                            });

                            // Log do erro
                            await Task.Run(() => sendEmail(ex.ToString(), Settings.Default.emailenvio, true, "informatica", "")).ConfigureAwait(false);
                        }
                    }
                }

                // Envia o relatório final
                await Task.Run(() => sendEmailRelatorio()).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                // Log do erro principal
                await Task.Run(() => sendEmail(ex.ToString(), Settings.Default.emailenvio, true, "informatica", "")).ConfigureAwait(false);
            }
        }

        private void sendEmailRelatorio()
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
                        if (listAcoesPublicadas != null && listAcoesPublicadas.Count > 0)
                        {
                            StringBuilder relatorio = new StringBuilder();
                            relatorio.AppendLine("Relatório de Publicação de Ações:<br><br>");

                            // Agrupa por tipo de ação
                            var acoesPublicadas = listAcoesPublicadas.Where(x => x.Acao.Contains("Publicada")).ToList();
                            var acoesNaoPublicadas = listAcoesPublicadas.Where(x => x.Acao.Contains("Não precisou")).ToList();
                            var acoesNaoProcessadas = listAcoesPublicadas.Where(x => x.Acao.Contains("Não processada")).ToList();
                            var acoesComErro = listAcoesPublicadas.Where(x => !x.Sucesso).ToList();

                            if (acoesPublicadas.Count > 0)
                            {
                                relatorio.AppendLine("<h3>Ações Publicadas:</h3>");
                                foreach (var acao in acoesPublicadas)
                                {
                                    relatorio.AppendLine($"<b>Ref Ação:</b> {acao.RefAccao} | <b>Curso:</b> {acao.Descricao} | <b>Ação:</b> {acao.Acao} | <b>Motivo:</b> {acao.Motivo}<br>");
                                }
                                relatorio.AppendLine("<br>");
                            }

                            if (acoesNaoPublicadas.Count > 0)
                            {
                                relatorio.AppendLine("<h3>Ações que não precisaram ser publicadas:</h3>");
                                foreach (var acao in acoesNaoPublicadas)
                                {
                                    relatorio.AppendLine($"<b>Ref Ação:</b> {acao.RefAccao} | <b>Curso:</b> {acao.Descricao} | <b>Motivo:</b> {acao.Motivo}<br>");
                                }
                                relatorio.AppendLine("<br>");
                            }

                            if (acoesNaoProcessadas.Count > 0)
                            {
                                relatorio.AppendLine("<h3>Ações não processadas:</h3>");
                                foreach (var acao in acoesNaoProcessadas)
                                {
                                    relatorio.AppendLine($"<b>Ref Ação:</b> {acao.RefAccao} | <b>Curso:</b> {acao.Descricao} | <b>Motivo:</b> {acao.Motivo}<br>");
                                }
                                relatorio.AppendLine("<br>");
                            }

                            if (acoesComErro.Count > 0)
                            {
                                relatorio.AppendLine("<h3>Ações com erro:</h3>");
                                foreach (var acao in acoesComErro)
                                {
                                    relatorio.AppendLine($"<b>Ref Ação:</b> {acao.RefAccao} | <b>Curso:</b> {acao.Descricao} | <b>Erro:</b> {acao.Erro}<br>");
                                }
                                relatorio.AppendLine("<br>");
                            }

                            body = relatorio.ToString();
                        }
                        else
                        {
                            body = "Não há ações para processar no dia de hoje.";
                        }

                        mm.Subject = "Instituto CRIAP || Relatório - Publicação de Ações " + data;
                        mm.BodyEncoding = UTF8Encoding.UTF8;
                        mm.IsBodyHtml = true;
                        mm.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure;
                        mm.Body = body + "<br> " + versao;

                        client.Send(mm);
                    }
                }
            }
            catch (Exception ex)
            {
                sendEmail(ex.ToString(), Settings.Default.emailenvio, true, "informatica", "");
            }
        }

        private void sendEmail(string body, string tecnica = "", bool error = false, string emailPessoa = "", string temp = "")
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

                mm.Subject = (!teste) ? "Publicação de Ações / " : data + " TESTE - Publicação de Ações // ";
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

        public void RegistraLog(string refAcao, string mensagem, string menu, string refAcaoLog)
        {
            List<objLogSendersFormador> logSenders = new List<objLogSendersFormador>();
            logSenders.Add(new objLogSendersFormador
            {
                idFormador = refAcao,
                mensagem = mensagem,
                menu = menu,
                refAccao = refAcaoLog
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