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
        public DateTime data;
        public string versao;
        private List<AcaoParaPublicacao> acoesParaProcessar;
        List<RelatorioPublicacao> listAcoesPublicadas = new List<RelatorioPublicacao>();

        // URLs das APIs - Updated to use POST format
        private const string API_BASE_URL = "http://localhost:5141/api/PublicaAcaoWeb";

        public Form1()
        {
            InitializeComponent();
            Security.remote();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            teste = true;

            if(teste)
                data = new DateTime(2025,10,12);
            else
                data = DateTime.Now.AddDays(-1);

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

                // Format the date for SQL
                string dataSql = data.ToString("yyyy-MM-dd");

                // Primeiro, verifica ações que precisam ser despublicadas após 24h
                string queryDespublicacao24h = $@"
                    DECLARE @data_base DATE = CAST('{dataSql}' AS DATE);
                    
                    -- Busca ações publicadas há mais de 24 horas que ainda não foram despublicadas
                    SELECT DISTINCT 
                        sp.RefAcaoParaDespublicar,
                        sp.DataQuandoPublicouSubstituta,
                        sp.MotivoSubstituicao,
                        ISNULL(c.Descricao, sp.RefAcaoParaDespublicar) AS Descricao -- Busca nome real do curso
                    FROM secretariaVirtual.dbo.sv_rotina_publicacoes_web_24h sp
                    INNER JOIN humantrain.dbo.TBForAccoes a ON a.Ref_Accao = sp.RefAcaoParaDespublicar
                    INNER JOIN humantrain.dbo.TBForCursos c ON c.Codigo_Curso = a.Codigo_Curso
                    WHERE sp.DataQuandoDespublicouOriginal IS NULL
                    AND DATEDIFF(hour, sp.DataQuandoPublicouSubstituta, '{data:yyyy-MM-dd HH:mm:ss}') >= 24
                    ORDER BY sp.DataQuandoPublicouSubstituta;";

                // Executa despublicação de ações pendentes há mais de 24h
                await Task.Run(() =>
                {
                    DataTable dataTableDespublicacao = new DataTable();
                    try
                    {
                        Connect.SVlocalConnect.ConnInit();
                        using (SqlDataAdapter adapterDespublicacao = new SqlDataAdapter(queryDespublicacao24h, Connect.SVlocalConnect.Conn))
                        {
                            adapterDespublicacao.Fill(dataTableDespublicacao);
                        }
                        Connect.SVlocalConnect.ConnEnd();

                        // Processa despublicações pendentes
                        if (dataTableDespublicacao.Rows.Count > 0)
                        {
                            using (var httpClient = new HttpClient())
                            {
                                httpClient.Timeout = TimeSpan.FromSeconds(30);

                                foreach (DataRow row in dataTableDespublicacao.Rows)
                                {
                                    string refAccao = row["RefAcaoParaDespublicar"]?.ToString()?.Trim();
                                    string dataPublicacao = row["DataQuandoPublicouSubstituta"]?.ToString();
                                    string motivo = row["MotivoSubstituicao"]?.ToString()?.Trim();
                                    string descricao = row["Descricao"]?.ToString()?.Trim();

                                    try
                                    {
                                        string apiUrlDespublicar = $"{API_BASE_URL}/despublicar/{refAccao}";
                                        var response = httpClient.PostAsync(apiUrlDespublicar, null).Result;
                                        
                                        bool sucessoDespublicacao = response.IsSuccessStatusCode;
                                        string mensagem = response.Content.ReadAsStringAsync().Result;

                                        // Marca como despublicada na tabela de controle - usar conexão SV
                                        Connect.SVlocalConnect.ConnInit();
                                        string updateQuery = $@"
                                            UPDATE sv_rotina_publicacoes_web_24h 
                                            SET DataQuandoDespublicouOriginal = '{data:yyyy-MM-dd HH:mm:ss}', 
                                                StatusDespublicacao = '{(sucessoDespublicacao ? "Sucesso" : "Erro")}',
                                                MensagemDespublicacao = '{mensagem.Replace("'", "''")}'
                                            WHERE RefAcaoParaDespublicar = '{refAccao}' AND DataQuandoDespublicouOriginal IS NULL";
                                        
                                        using (SqlCommand cmdUpdate = new SqlCommand(updateQuery, Connect.SVlocalConnect.Conn))
                                        {
                                            cmdUpdate.ExecuteNonQuery();
                                        }
                                        Connect.SVlocalConnect.ConnEnd();

                                        if (sucessoDespublicacao)
                                        {
                                            // Verifica se já estava despublicada - tratar como "não precisou"
                                            if (mensagem.ToLower().Contains("já está despublicada") || mensagem.ToLower().Contains("already"))
                                            {
                                                listAcoesPublicadas.Add(new RelatorioPublicacao
                                                {
                                                    RefAccao = refAccao,
                                                    Descricao = descricao,
                                                    Acao = "Não precisou despublicar",
                                                    Motivo = $"Ação já estava despublicada - Publicada em {dataPublicacao} - Motivo original: {motivo}",
                                                    PercAceitacao = 0,
                                                    Sucesso = true,
                                                    Erro = ""
                                                });
                                            }
                                            else
                                            {
                                                listAcoesPublicadas.Add(new RelatorioPublicacao
                                                {
                                                    RefAccao = refAccao,
                                                    Descricao = descricao,
                                                    Acao = "Despublicada após 24h",
                                                    Motivo = $"Publicada em {dataPublicacao} - Motivo original: {motivo}",
                                                    PercAceitacao = 0,
                                                    Sucesso = true,
                                                    Erro = ""
                                                });
                                            }

                                            RegistraLog(refAccao, $"Ação {refAccao} despublicada automaticamente após 24h - Publicada em {dataPublicacao}", "Despublicação Automática", refAccao);
                                        }
                                        else
                                        {
                                            listAcoesPublicadas.Add(new RelatorioPublicacao
                                            {
                                                RefAccao = refAccao,
                                                Descricao = descricao,
                                                Acao = "Erro na despublicação após 24h",
                                                Motivo = $"Publicada em {dataPublicacao} - Motivo original: {motivo}",
                                                PercAceitacao = 0,
                                                Sucesso = false,
                                                Erro = mensagem
                                            });
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        listAcoesPublicadas.Add(new RelatorioPublicacao
                                        {
                                            RefAccao = refAccao,
                                            Descricao = descricao,
                                            Acao = "Erro na despublicação após 24h",
                                            Motivo = $"Publicada em {dataPublicacao} - Motivo original: {motivo}",
                                            PercAceitacao = 0,
                                            Sucesso = false,
                                            Erro = ex.Message
                                        });
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // Log do erro de despublicação
                        sendEmail($"Erro na rotina de despublicação após 24h: {ex.Message}", Settings.Default.emailenvio, true, "informatica", "");
                    }
                }).ConfigureAwait(false);

                // Busca as ações que precisam ser processadas
                string queryAcoes = $@"
                    DECLARE @data_base DATE = CAST('{dataSql}' AS DATE); -- Data de hoje

                    -- ======================================================================
                    -- Blocos de estados auditados
                    -- ======================================================================
                    WITH adiados AS (
	                    SELECT versao_rowid_Accao, versao_data
	                    FROM (
		                    SELECT *,
			                       ROW_NUMBER() OVER (PARTITION BY versao_rowid_Accao ORDER BY versao_data DESC) AS rn
		                    FROM TBAuditAccoes
		                    WHERE Codigo_Estado = 2 AND Codigo_Estado_Antigo != 2
                            AND DATEDIFF(day, CAST(versao_data AS DATE), @data_base) = 3 -- Exatamente 3 dias após mudança para adiado
	                    ) o
	                    WHERE rn = 1
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
	                    -- Fases caducadas: verifica se a última fase de candidatura termina hoje
	                    (c.Tipo_Curso IN ('11','12') AND CAST(COALESCE(Qfasecandidatura.qfasedata, TfaseCandidatura.tfasedata, SfaseCandidatura.sfasedata, Pfasecandidatura.pfasedata) AS DATE) = @data_base)
                     OR
	                    (c.Tipo_Curso NOT IN ('11','12') AND CAST(COALESCE(Qfasecandidatura.qfasedata, TfaseCandidatura.tfasedata, SfaseCandidatura.sfasedata, Pfasecandidatura.pfasedata) AS DATE) = @data_base)
                     OR
	                    -- Ações adiadas: exatamente 3 dias após mudança de estado
	                    (adiados.versao_rowid_Accao IS NOT NULL)
                     OR
	                    -- Ações confirmadas: processadas no mesmo dia da mudança
	                    (confirmados.versao_rowid_Accao IS NOT NULL)
                     OR
	                    -- Ações confirmadas provisórias: processadas no mesmo dia da mudança
	                    (confirmados_provisorios.versao_rowid_Accao IS NOT NULL)
                     OR
	                    -- Ações canceladas: processadas no mesmo dia da mudança
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
                    DataMotivo = row.IsNull("Data_Motivo") ? null : row["Data_Motivo"]?.ToString(),
                    StatusTratamento = row.Field<string>("Status_Tratamento")?.Trim(),
                    ProxDataInicio = row.Field<string>("Prox_Data_Inicio")?.Trim(),
                    ProxRefAccao = row.Field<string>("Prox_Ref_Accao")?.Trim(),
                    ProxEstado = row.Field<string>("Prox_Estado")?.Trim(),
                    ProxPubWeb = row.IsNull("Prox_Pub_Web") ? null : row["Prox_Pub_Web"]?.ToString()
                }).ToList();

                await Task.Run(() => Connect.HTlocalConnect.ConnEnd()).ConfigureAwait(false);

                // Processa cada ação
                using (var httpClient = new HttpClient())
                {
                    httpClient.Timeout = TimeSpan.FromSeconds(30);

                    foreach (var acao in acoesParaProcessar)
                    {
                        if (acao.RefAccao.ToUpper().StartsWith("FM_") || acao.RefAccao.ToUpper().Contains("CFPIF"))
                        {
                            continue;
                        }

                        // Processa apenas ações com Status_Tratamento = "Já Tratado" e Motivo = "Fases Caducadas" ou "Adiado"
                        bool deveProcessar = acao.StatusTratamento == "Já Tratado" && 
                                           (acao.Motivo == "Fases Caducadas" || acao.Motivo == "Adiado");
                        
                        if (!deveProcessar)
                        {
                            // Adiciona ao relatório como "não processada" com classificação específica
                            string acaoDescricao;
                            if (acao.Motivo == "Confirmado")
                                acaoDescricao = "Não processada - Confirmada";
                            else if (acao.Motivo == "Confirmado Provisório")
                                acaoDescricao = "Não processada - Confirmada Provisória";
                            else if (acao.Motivo == "Cancelado")
                                acaoDescricao = "Não processada - Cancelada";
                            else if (acao.StatusTratamento != "Já Tratado")
                                acaoDescricao = "Não processada - Não Tratada";
                            else
                                acaoDescricao = "Não processada - Outro Motivo";
                                
                            listAcoesPublicadas.Add(new RelatorioPublicacao
                            {
                                RefAccao = acao.RefAccao,
                                Descricao = acao.Descricao,
                                Acao = acaoDescricao,
                                Motivo = $"Status: {acao.StatusTratamento} | Motivo: {acao.Motivo}",
                                PercAceitacao = acao.PercAceitacao,
                                Sucesso = true,
                                Erro = ""
                            });
                            continue;
                        }

                        // Para ações processáveis com mais de 50% de aceitação - não faz nada
                        if (acao.PercAceitacao >= 50)
                        {
                            listAcoesPublicadas.Add(new RelatorioPublicacao
                            {
                                RefAccao = acao.RefAccao,
                                Descricao = acao.Descricao,
                                Acao = "Não precisou publicar",
                                Motivo = $"Aceitação >= 50% ({acao.PercAceitacao}%) | Status: {acao.StatusTratamento} | Motivo: {acao.Motivo}",
                                PercAceitacao = acao.PercAceitacao,
                                Sucesso = true,
                                Erro = ""
                            });

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
                                Motivo = $"Aceitação < 50% ({acao.PercAceitacao}%) mas não há próxima ação | Status: {acao.StatusTratamento} | Motivo: {acao.Motivo}",
                                PercAceitacao = acao.PercAceitacao,
                                Sucesso = false,
                                Erro = "Próxima ação não encontrada"
                            });
                            continue;
                        }

                        // Chama a API de publicação usando POST request
                        try
                        {
                            bool publicacaoSucesso = false;
                            string mensagemPublicacao = "";
                            string erroGeral = "";

                            // Publica APENAS a próxima ação (não despublica a original)
                            // A ação original (adiada/caducada) permanece despublicada
                            try
                            {
                                string apiUrlPublicar = $"{API_BASE_URL}/publicar/{acao.ProxRefAccao}";
                                var responsePublicar = await httpClient.PostAsync(apiUrlPublicar, null).ConfigureAwait(false);
                                var responseContentPublicar = await responsePublicar.Content.ReadAsStringAsync().ConfigureAwait(false);

                                if (responsePublicar.IsSuccessStatusCode)
                                {
                                    try
                                    {
                                        var publicacaoResponse = JsonConvert.DeserializeObject<PublicacaoResponse>(responseContentPublicar);
                                        publicacaoSucesso = publicacaoResponse?.Sucesso ?? true;
                                        mensagemPublicacao = publicacaoResponse?.Mensagem ?? "";
                                    }
                                    catch
                                    {
                                        publicacaoSucesso = true;
                                        mensagemPublicacao = responseContentPublicar;
                                    }
                                }
                                else
                                {
                                    // Tenta extrair erro estruturado
                                    try
                                    {
                                        var errorResponse = JsonConvert.DeserializeObject<ErrorResponse>(responseContentPublicar);
                                        mensagemPublicacao = errorResponse?.MensagemErro ?? responseContentPublicar;
                                    }
                                    catch
                                    {
                                        mensagemPublicacao = responseContentPublicar;
                                    }
                                    
                                    // Verifica se é um erro "já publicada" - tratar como sucesso
                                    if (mensagemPublicacao.ToLower().Contains("já está publicada") || 
                                        responsePublicar.StatusCode == System.Net.HttpStatusCode.Conflict)
                                    {
                                        publicacaoSucesso = true; // Tratar como sucesso
                                    }
                                    else
                                    {
                                        erroGeral += $"Publicação falhou (HTTP {responsePublicar.StatusCode}): {mensagemPublicacao}. ";
                                    }
                                }
                            }
                            catch (Exception exPublicar)
                            {
                                mensagemPublicacao = exPublicar.Message;
                                erroGeral += $"Erro na publicação: {mensagemPublicacao}. ";
                            }

                            // Avalia o resultado geral - agora só verifica a publicação
                            bool sucessoGeral = publicacaoSucesso;
                            
                            // Verifica se a próxima ação já estava publicada
                            bool jaEstavaPublicada = mensagemPublicacao.ToLower().Contains("já está publicada");

                            if (sucessoGeral)
                            {
                                // Se a próxima ação já estava publicada, classificar como "não precisou publicar"
                                if (jaEstavaPublicada)
                                {
                                    listAcoesPublicadas.Add(new RelatorioPublicacao
                                    {
                                        RefAccao = acao.RefAccao,
                                        Descricao = acao.Descricao,
                                        Acao = "Não precisou publicar",
                                        Motivo = $"Próxima ação {acao.ProxRefAccao} já estava publicada | Aceitação < 50% ({acao.PercAceitacao}%) | Status: {acao.StatusTratamento} | Motivo: {acao.Motivo}",
                                        PercAceitacao = acao.PercAceitacao,
                                        Sucesso = true,
                                        Erro = ""
                                    });
                                }
                                else
                                {
                                    string mensagemCompleta = $"Publicada próxima ação: {acao.ProxRefAccao}";
                                    if (!string.IsNullOrEmpty(mensagemPublicacao))
                                    {
                                        mensagemCompleta += $" ({mensagemPublicacao})";
                                    }

                                    // CORREÇÃO: Agenda despublicação da AÇÃO ORIGINAL (não da próxima que foi publicada)
                                    // Lógica correta: publica próxima ação Y → agenda despublicação da ação original X após 24h
                                    await Task.Run(() =>
                                {
                                    try
                                    {
                                        Connect.SVlocalConnect.ConnInit(); // Usar conexão SV onde está a tabela
                                        
                                        // Verifica se a tabela existe, se não existir cria
                                        string checkTableQuery = @"
                                            IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='sv_rotina_publicacoes_web_24h' AND xtype='U')
                                            BEGIN
                                                CREATE TABLE sv_rotina_publicacoes_web_24h (
                                                    Id INT IDENTITY(1,1) PRIMARY KEY,
                                                    RefAcaoParaDespublicar VARCHAR(50) NOT NULL,
                                                    DataQuandoPublicouSubstituta DATETIME NOT NULL,
                                                    MotivoSubstituicao VARCHAR(500),
                                                    RefAcaoSubstitutaPublicada VARCHAR(50),
                                                    DataQuandoDespublicouOriginal DATETIME NULL,
                                                    StatusDespublicacao VARCHAR(20) NULL,
                                                    MensagemDespublicacao VARCHAR(500) NULL,
                                                    DataCriacao DATETIME DEFAULT GETDATE()
                                                );
                                            END";
                                        
                                        using (SqlCommand cmdCheck = new SqlCommand(checkTableQuery, Connect.SVlocalConnect.Conn))
                                        {
                                            cmdCheck.ExecuteNonQuery();
                                        }

                                        // Agenda a despublicação da AÇÃO ORIGINAL (não da próxima)
                                        string insertControlQuery = $@"
                                            INSERT INTO sv_rotina_publicacoes_web_24h 
                                            (RefAcaoParaDespublicar, DataQuandoPublicouSubstituta, MotivoSubstituicao, RefAcaoSubstitutaPublicada) 
                                            VALUES 
                                            ('{acao.RefAccao}', '{data:yyyy-MM-dd HH:mm:ss}', '{acao.Motivo.Replace("'", "''")} - Aceitação {acao.PercAceitacao}% - Próxima ação publicada: {acao.ProxRefAccao}', '{acao.ProxRefAccao}')";
                                        
                                        using (SqlCommand cmdInsert = new SqlCommand(insertControlQuery, Connect.SVlocalConnect.Conn))
                                        {
                                            cmdInsert.ExecuteNonQuery();
                                        }
                                        Connect.SVlocalConnect.ConnEnd();
                                    }
                                    catch (Exception ex)
                                    {
                                        // Log do erro de inserção na tabela de controle
                                        sendEmail($"Erro ao registrar publicação na tabela de controle: {ex.Message}", Settings.Default.emailenvio, true, "informatica", "");
                                    }
                                    }).ConfigureAwait(false);

                                    listAcoesPublicadas.Add(new RelatorioPublicacao
                                    {
                                        RefAccao = acao.RefAccao,
                                        Descricao = acao.Descricao,
                                        Acao = mensagemCompleta,
                                        Motivo = $"Aceitação < 50% ({acao.PercAceitacao}%) | Status: {acao.StatusTratamento} | Motivo: {acao.Motivo}",
                                        PercAceitacao = acao.PercAceitacao,
                                        Sucesso = true,
                                        Erro = ""
                                    });

                                    // Registra log
                                    await Task.Run(() => RegistraLog(acao.RefAccao, $"Despublicada ação {acao.RefAccao} e publicada próxima ação {acao.ProxRefAccao} - Aceitação < 50% ({acao.PercAceitacao}%) - Status: {acao.StatusTratamento} - Motivo: {acao.Motivo}", "Publicação Ação", acao.RefAccao)).ConfigureAwait(false);
                                }
                            }
                            else
                            {
                                listAcoesPublicadas.Add(new RelatorioPublicacao
                                {
                                    RefAccao = acao.RefAccao,
                                    Descricao = acao.Descricao,
                                    Acao = "Erro na operação",
                                    Motivo = $"Aceitação < 50% ({acao.PercAceitacao}%) | Status: {acao.StatusTratamento} | Motivo: {acao.Motivo}",
                                    PercAceitacao = acao.PercAceitacao,
                                    Sucesso = false,
                                    Erro = erroGeral.Trim()
                                });
                            }
                        }
                        catch (Exception ex)
                        {
                            listAcoesPublicadas.Add(new RelatorioPublicacao
                            {
                                RefAccao = acao.RefAccao,
                                Descricao = acao.Descricao,
                                Acao = "Erro na operação",
                                Motivo = $"Aceitação < 50% ({acao.PercAceitacao}%) | Status: {acao.StatusTratamento} | Motivo: {acao.Motivo}",
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
                            mm.To.Add("planeamento@criap.com");
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
                            var acoesPublicadas = listAcoesPublicadas.Where(x => x.Acao.Contains("Publicada") && !x.Acao.Contains("Despublicada")).ToList();
                            var acoesDespublicadas = listAcoesPublicadas.Where(x => x.Acao.Contains("Despublicada")).ToList();
                            var acoesNaoPublicadas = listAcoesPublicadas.Where(x => x.Acao.Contains("Não precisou")).ToList();
                            
                            // Separa ações não processadas por motivo específico
                            var acoesConfirmadas = listAcoesPublicadas.Where(x => x.Acao.Contains("Confirmada") && !x.Acao.Contains("Provisória")).ToList();
                            var acoesConfirmadasProvisionais = listAcoesPublicadas.Where(x => x.Acao.Contains("Confirmada Provisória")).ToList();
                            var acoesCanceladas = listAcoesPublicadas.Where(x => x.Acao.Contains("Cancelada")).ToList();
                            var acoesNaoTratadas = listAcoesPublicadas.Where(x => x.Acao.Contains("Não Tratada")).ToList();
                            var acoesOutroMotivo = listAcoesPublicadas.Where(x => x.Acao.Contains("Outro Motivo")).ToList();
                            
                            var acoesComErro = listAcoesPublicadas.Where(x => !x.Sucesso).ToList();

                            // Verifica se houve processamento
                            bool houveProcessamento = (acoesPublicadas.Count > 0 || acoesDespublicadas.Count > 0 || acoesNaoPublicadas.Count > 0);

                            if (!houveProcessamento)
                            {
                                relatorio.AppendLine("<h3>Nenhuma ação foi processada hoje.</h3>");
                                relatorio.AppendLine("Não foram encontradas ações que necessitem de intervenção automática.<br><br>");
                                relatorio.AppendLine("<br><hr><br>");
                            }
                            else
                            {
                                if (acoesPublicadas.Count > 0)
                                {
                                    relatorio.AppendLine("<h3>Ações Publicadas:</h3>");
                                    relatorio.AppendLine("<table border='1' cellpadding='8' cellspacing='0' style='border-collapse: collapse; width: 100%;'>");
                                    relatorio.AppendLine("<tr style='background-color: #f0f0f0; font-weight: bold;'>");
                                    relatorio.AppendLine("<th>Ref Ação</th><th>Curso</th><th>Ação Realizada</th><th>% Aceitação</th><th>Status/Motivo</th>");
                                    relatorio.AppendLine("</tr>");
                                    
                                    foreach (var acao in acoesPublicadas)
                                    {
                                        // Extrai informações da ação para melhor apresentação
                                        string acaoSimplificada = acao.Acao;
                                        if (acao.Acao.Contains("Publicada próxima ação:"))
                                        {
                                            string proximaAcao = acao.Acao.Replace("Publicada próxima ação:", "").Trim();
                                            // Remove texto entre parênteses se existir
                                            if (proximaAcao.Contains("(")) proximaAcao = proximaAcao.Substring(0, proximaAcao.IndexOf("(")).Trim();
                                            
                                            acaoSimplificada = $"<span style='color: #008000;'>Publicou:</span> {proximaAcao}";
                                        }
                                        else if (acao.Acao.Contains("Despublicada após 24h"))
                                        {
                                            acaoSimplificada = "<span style='color: #CC6600;'>Despublicada após 24h</span>";
                                        }
                                        
                                        relatorio.AppendLine("<tr>");
                                        relatorio.AppendLine($"<td><b>{acao.RefAccao}</b></td>");
                                        relatorio.AppendLine($"<td>{acao.Descricao}</td>");
                                        relatorio.AppendLine($"<td>{acaoSimplificada}</td>");
                                        relatorio.AppendLine($"<td style='text-align: center;'>{acao.PercAceitacao:F2}%</td>");
                                        relatorio.AppendLine($"<td>{acao.Motivo}</td>");
                                        relatorio.AppendLine("</tr>");
                                    }
                                    relatorio.AppendLine("</table><br>");
                                }

                                if (acoesDespublicadas.Count > 0)
                                {
                                    relatorio.AppendLine("<h3>Ações Despublicadas:</h3>");
                                    relatorio.AppendLine("<table border='1' cellpadding='8' cellspacing='0' style='border-collapse: collapse; width: 100%;'>");
                                    relatorio.AppendLine("<tr style='background-color: #ffe6e6; font-weight: bold;'>");
                                    relatorio.AppendLine("<th>Ref Ação</th><th>Curso</th><th>Ação Realizada</th><th>Motivo</th>");
                                    relatorio.AppendLine("</tr>");
                                    
                                    foreach (var acao in acoesDespublicadas)
                                    {
                                        // Extrai informações da ação para melhor apresentação
                                        string acaoSimplificada = acao.Acao;
                                        if (acao.Acao.Contains("Despublicada após 24h"))
                                        {
                                            acaoSimplificada = "<span style='color: #CC6600;'>Despublicada após 24h</span>";
                                        }
                                        else if (acao.Acao.Contains("Não precisou despublicar"))
                                        {
                                            acaoSimplificada = "<span style='color: #666666;'>Já estava despublicada</span>";
                                        }
                                        else if (acao.Acao.Contains("Erro na despublicação"))
                                        {
                                            acaoSimplificada = "<span style='color: #FF0000;'>Erro na despublicação</span>";
                                        }
                                        
                                        relatorio.AppendLine("<tr>");
                                        relatorio.AppendLine($"<td><b>{acao.RefAccao}</b></td>");
                                        relatorio.AppendLine($"<td>{acao.Descricao}</td>");
                                        relatorio.AppendLine($"<td>{acaoSimplificada}</td>");
                                        relatorio.AppendLine($"<td>{acao.Motivo}</td>");
                                        relatorio.AppendLine("</tr>");
                                    }
                                    relatorio.AppendLine("</table><br>");
                                }

                                if (acoesNaoPublicadas.Count > 0)
                                {
                                    relatorio.AppendLine("<h3>Ações que não precisaram ser publicadas:</h3>");
                                    relatorio.AppendLine("<table border='1' cellpadding='6' cellspacing='0' style='border-collapse: collapse; width: 100%;'>");
                                    relatorio.AppendLine("<tr style='background-color: #f8f8f8; font-weight: bold;'>");
                                    relatorio.AppendLine("<th>Ref Ação</th><th>Curso</th><th>% Aceitação</th><th>Motivo</th>");
                                    relatorio.AppendLine("</tr>");

                                    foreach (var acao in acoesNaoPublicadas)
                                    {
                                        relatorio.AppendLine("<tr>");
                                        relatorio.AppendLine($"<td><b>{acao.RefAccao}</b></td>");
                                        relatorio.AppendLine($"<td>{acao.Descricao}</td>");
                                        relatorio.AppendLine($"<td style='text-align: center;'>{acao.PercAceitacao:F2}%</td>");
                                        relatorio.AppendLine($"<td>{acao.Motivo}</td>");
                                        relatorio.AppendLine("</tr>");
                                    }
                                    relatorio.AppendLine("</table><br>");
                                }
                                relatorio.AppendLine("<br><hr><br>");
                            }

                            // Inicia seção de informações adicionais apenas se houver conteúdo
                            bool temInformacoesAdicionais = (acoesConfirmadas.Count > 0 || acoesConfirmadasProvisionais.Count > 0 || 
                                                           acoesCanceladas.Count > 0 || acoesNaoTratadas.Count > 0 || 
                                                           acoesOutroMotivo.Count > 0 || acoesComErro.Count > 0);

                            if (temInformacoesAdicionais)
                            {
                                relatorio.AppendLine("<h3>Informações Adicionais:</h3>");
                                relatorio.AppendLine("<div style='font-size: 90%;'>");

                                if (acoesConfirmadas.Count > 0)
                                {
                                    relatorio.AppendLine("<b>Ações confirmadas (não processadas):</b><br>");
                                    foreach (var acao in acoesConfirmadas)
                                    {
                                        relatorio.AppendLine($"Ref Ação: {acao.RefAccao} | Curso: {acao.Descricao} | % Aceitação: {acao.PercAceitacao}%<br>");
                                    }
                                    relatorio.AppendLine("<br>");
                                }

                                if (acoesConfirmadasProvisionais.Count > 0)
                                {
                                    relatorio.AppendLine("<b>Ações confirmadas provisórias (não processadas):</b><br>");
                                    foreach (var acao in acoesConfirmadasProvisionais)
                                    {
                                        relatorio.AppendLine($"Ref Ação: {acao.RefAccao} | Curso: {acao.Descricao} | % Aceitação: {acao.PercAceitacao}%<br>");
                                    }
                                    relatorio.AppendLine("<br>");
                                }

                                if (acoesCanceladas.Count > 0)
                                {
                                    relatorio.AppendLine("<b>Ações canceladas (não processadas):</b><br>");
                                    foreach (var acao in acoesCanceladas)
                                    {
                                        relatorio.AppendLine($"Ref Ação: {acao.RefAccao} | Curso: {acao.Descricao} | % Aceitação: {acao.PercAceitacao}%<br>");
                                    }
                                    relatorio.AppendLine("<br>");
                                }

                                if (acoesNaoTratadas.Count > 0)
                                {
                                    relatorio.AppendLine("<b>Ações não tratadas (não processadas):</b><br>");
                                    foreach (var acao in acoesNaoTratadas)
                                    {
                                        // Aplica cor vermelha para "Não Tratado" e verde para "Já Publicado"
                                        string motivoComCor = acao.Motivo;
                                        if (acao.Motivo.Contains("Não Tratado"))
                                        {
                                            motivoComCor = acao.Motivo.Replace("Não Tratado", "<span style='color: #DC143C;'>Não Tratado</span>");
                                        }
                                        else if (acao.Motivo.Contains("Já Publicado"))
                                        {
                                            motivoComCor = acao.Motivo.Replace("Já Publicado", "<span style='color: #008000;'>Já Publicado</span>");
                                        }
                                        
                                        relatorio.AppendLine($"Ref Ação: {acao.RefAccao} | Curso: {acao.Descricao} | Status: {motivoComCor}<br>");
                                    }
                                    relatorio.AppendLine("<br>");
                                }

                                if (acoesOutroMotivo.Count > 0)
                                {
                                    relatorio.AppendLine("<b>Outras ações não processadas:</b><br>");
                                    foreach (var acao in acoesOutroMotivo)
                                    {
                                        // Aplica cor vermelha para "Não Tratado" e verde para "Já Publicado"
                                        string motivoComCor = acao.Motivo;
                                        if (acao.Motivo.Contains("Não Tratado"))
                                        {
                                            motivoComCor = acao.Motivo.Replace("Não Tratado", "<span style='color: #DC143C;'>Não Tratado</span>");
                                        }
                                        else if (acao.Motivo.Contains("Já Publicado"))
                                        {
                                            motivoComCor = acao.Motivo.Replace("Já Publicado", "<span style='color: #008000;'>Já Publicado</span>");
                                        }
                                        
                                        relatorio.AppendLine($"Ref Ação: {acao.RefAccao} | Curso: {acao.Descricao} | Motivo: {motivoComCor}<br>");
                                    }
                                    relatorio.AppendLine("<br>");
                                }

                                if (acoesComErro.Count > 0)
                                {
                                    relatorio.AppendLine("<b>Ações com erro:</b><br>");
                                    foreach (var acao in acoesComErro)
                                    {
                                        relatorio.AppendLine($"Ref Ação: {acao.RefAccao} | Curso: {acao.Descricao} | Erro: {acao.Erro}<br>");
                                    }
                                    relatorio.AppendLine("<br>");
                                }

                                relatorio.AppendLine("</div>");
                            }

                            body = relatorio.ToString();
                        }
                        else
                        {
                            body = "Não há ações para processar no dia de hoje.";
                        }

                        mm.Subject = "Instituto CRIAP || Relatório - Publicação de Ações " + data.ToString("dd/MM/yyyy");
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
                        mm.To.Add("planeamento@criap.com");
                        mm.CC.Add("informatica@criap.com");
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

                mm.Subject = (!teste) ? "Publicação de Ações / " : data.ToString("dd/MM/yyyy") + " TESTE - Publicação de Ações // ";
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
                idFormador = "0", // Usar ID genérico para sistema, pois idFormando é INT
                mensagem = mensagem,
                menu = menu,
                refAccao = refAcaoLog
            });
            DataBaseLogSave(logSenders, data);
        }

        public static void DataBaseLogSave(List<objLogSendersFormador> logSenders, DateTime dataBase)
        {
            if (logSenders.Count > 0)
            {
                string dataFormatada = dataBase.ToString("yyyy-MM-dd HH:mm:ss");
                string subQuery = "INSERT INTO sv_logs (idFormando, refAcao, dataregisto, registo, menu, username) VALUES ";
                for (int i = 0; i < logSenders.Count; i++)
                {
                    if (i < logSenders.Count - 1)
                        subQuery += "('" + logSenders[i].idFormador + "', '" + logSenders[i].refAccao + "', '" + dataFormatada + "', '" + logSenders[i].mensagem + "', '" + logSenders[i].menu + "', 'system_rotina'), ";
                    else subQuery += "('" + logSenders[i].idFormador + "', '" + logSenders[i].refAccao + "', '" + dataFormatada + "', '" + logSenders[i].mensagem + "', '" + logSenders[i].menu + "', 'system_rotina') ";
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