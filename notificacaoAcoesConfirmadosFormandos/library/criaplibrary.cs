using criapLibrary.types;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;

namespace criapLibrary
{
    public class Sv
    {
        public static List<objCursos> GetHT_data(bool all, string username, SqlConnection connection)
        {
            List<objCursos> objCursosList = new List<objCursos>();

            string str1 = "SELECT a.Codigo_Curso, c.Descricao, CAST(a.Data_Inicio AS DATE) AS Data_Inicio, CAST(a.Data_Fim AS DATE) AS Data_Fim, c.Duracao_Defeito, l.Local, c.Codigo_Area, a.Numero_Accao, a.Ref_Accao, c.Tipo_Curso, a.Numero_Projecto, a.Numero_Accao, a.Valor_Accao, a.Cod_Tecn_Resp, a.Cod_Tecn_Assist FROM TBForAccoes a INNER JOIN TBForCursos c ON a.Codigo_Curso = c.Codigo_Curso INNER JOIN TBForCandAccoes ca ON a.Numero_Projecto = ca.Numero_Cand AND a.Codigo_Curso = ca.Codigo_Curso AND a.Numero_Accao = ca.Numero_Accao INNER JOIN TBForLocais l ON ca.Codigo_Local = l.Codigo_Local WHERE (a.Codigo_Estado IN (1, 4, 2)) AND (a.Ref_Accao IS NOT NULL)";
            string str2 = (string)null;
            if (!all)
                str2 = "AND (a.Cod_Tecn_Resp = '" + username + "' OR a.Cod_Tecn_Assist = '" + username + "')";
            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(str1 + str2, connection);
            DataTable dataTable = new DataTable();
            sqlDataAdapter.Fill(dataTable);
            List<DataRow> list = dataTable.AsEnumerable().ToList<DataRow>();
            connection.Close();
            if ((uint)list.Count > 0U)
            {
                for (int index = 0; index < list.Count; ++index)
                {
                    objCursos objCursos = new objCursos()
                    {
                        Codigo_Curso = list[index][0].ToString(),
                        Descricao = list[index][1].ToString(),
                        Data_Inicio = list[index][2].ToString(),
                        Data_Fim = list[index][3].ToString(),
                        Duracao_Defeito = list[index][4].ToString(),
                        Local = list[index][5].ToString(),
                        Area = list[index][6].ToString(),
                        Edicao = list[index][7].ToString(),
                        Ref_Accao = list[index][8].ToString(),
                        TipoCurso = list[index][9].ToString(),
                        Numero_Projecto = list[index][10].ToString(),
                        Numero_Accao = list[index][11].ToString(),
                        Valor_Accao = float.Parse(list[index][12].ToString()).ToString(),
                        Cod_Tecn_Resp = list[index][13].ToString(),
                        Cod_Tecn_Assist = list[index][14].ToString()
                    };
                    char[] separator = new char[1] { '/' };
                    string[] strArray = objCursos.Numero_Projecto.Split(separator, StringSplitOptions.RemoveEmptyEntries);
                    string str3 = "\\\\192.168.1.248\\docs\\Projectos\\";
                    string str4;
                    if (((IEnumerable<string>)strArray).Count<string>() == 2)
                        str4 = str3 + strArray[0] + "\\" + strArray[1] + "\\" + objCursos.Numero_Accao + "." + objCursos.Codigo_Curso;
                    else
                        str4 = str3 + strArray[0] + "\\" + objCursos.Numero_Accao + "." + objCursos.Codigo_Curso;
                    objCursos.Pasta = str4;
                    objCursosList.Add(objCursos);
                }
            }
            return objCursosList;
        }
        public static List<library.types.objSalas> getHT_Salas(SqlConnection connection)
        {
            List<library.types.objSalas> objSalas = new List<library.types.objSalas>();
            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter("SELECT Codigo_Sala, Sala FROM TbForSalas ", connection);
            DataTable dataTable = new DataTable();
            sqlDataAdapter.Fill(dataTable);
            List<DataRow> list = dataTable.AsEnumerable().ToList<DataRow>();
            connection.Close();
            if ((uint)list.Count > 0U)
            {
                for (int index = 0; index < list.Count; ++index)
                {
                    library.types.objSalas objSala = new library.types.objSalas()
                    {
                        codigo_sala = list[index][0].ToString(),
                        sala = list[index][1].ToString()
                    };
                    objSalas.Add(objSala);
                }
            }
            return objSalas;
        }
        public static List<objColaboradores> getHT_Users(SqlConnection connection)
        {
            List<objColaboradores> objColaboradoresList = new List<objColaboradores>();
            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter("SELECT su.Login, su.Nome, su.Email, Sigla.valor AS Sigla, su.Telefone, Cor.valor AS Cor, su.versao_rowid, su.Conta_Activa, Foto.valor FROM TBSysUsers su LEFT OUTER JOIN (SELECT id_ecran, rowid_ecran, valor FROM TBGerCUValores AS TBGerCUValores_1 WHERE (nome_campo = 'zfoto') AND(id_ecran = 21)) AS Foto ON su.versao_rowid = Foto.rowid_ecran LEFT OUTER JOIN (SELECT id_ecran, rowid_ecran, valor FROM TBGerCUValores AS TBGerCUValores_1 WHERE(nome_campo = 'zcor') AND(id_ecran = 21)) AS Cor ON su.versao_rowid = Cor.rowid_ecran LEFT OUTER JOIN (SELECT id_ecran, rowid_ecran, valor FROM TBGerCUValores AS TBGerCUValores_1 WHERE(nome_campo = 'sigla') AND(id_ecran = 21)) AS Sigla ON su.versao_rowid = Sigla.rowid_ecran WHERE(su.Conta_Activa = 1) AND (su.Login NOT IN('hs', 'testes', 'ana.silva_1'))", connection);
            DataTable dataTable = new DataTable();
            sqlDataAdapter.Fill(dataTable);
            List<DataRow> list = dataTable.AsEnumerable().ToList<DataRow>();
            connection.Close();
            if ((uint)list.Count > 0U)
            {
                for (int index = 0; index < list.Count; ++index)
                {
                    objColaboradores objColaboradores = new objColaboradores()
                    {
                        username = list[index][0].ToString(),
                        nome = list[index][1].ToString(),
                        email = list[index][2].ToString(),
                        sigla = list[index][3].ToString(),
                        cor = ColorTranslator.FromHtml(list[index][5].ToString()),
                        colaboradorID = list[index][6].ToString(),
                        telefone = list[index][4].ToString()
                    };
                    if (list[index][8] != DBNull.Value)
                        objColaboradores.foto = Sv.stringToImage(list[index][8].ToString());
                    objColaboradoresList.Add(objColaboradores);
                }
            }
            return objColaboradoresList;
        }
        public static Image stringToImage(string inputString)
        {
            byte[] buffer = Convert.FromBase64String(inputString);
            MemoryStream memoryStream = new MemoryStream(buffer, 0, buffer.Length);
            memoryStream.Write(buffer, 0, buffer.Length);
            return Image.FromStream((Stream)memoryStream, true);
        }
        public static List<DataRow> getHTCursosSessoes(SqlConnection connection)
        {
            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter("SELECT a.Codigo_Curso, a.Ref_Accao, a.versao_rowid, c.Descricao, CAST(a.Data_Inicio AS DATE) AS Data_Inicio, CAST(a.Data_Fim AS DATE) AS Data_Fim, ISNULL(c.Duracao_Defeito, 0) as Duracao_Defeito FROM TBForAccoes a INNER JOIN TBForCursos c ON a.Codigo_Curso = c.Codigo_Curso WHERE a.Ref_Accao is NOT null AND (a.Codigo_Estado IN (1, 4, 2))", connection);
            DataTable dataTable = new DataTable();
            sqlDataAdapter.Fill(dataTable);
            List<DataRow> list = dataTable.AsEnumerable().ToList<DataRow>();
            connection.Close();
            return list;
        }
        public static List<DataRow> getHTAccoes(string rowidAccao, SqlConnection connection)
        {
            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter("SELECT s.Rowid_Accao, s.Rowid_Modulo, s.Num_Sessao, s.Data, s.Hora_Inicio, s.Hora_Fim, s.Codigo_Formador, s.SC, s.CT, s.PS, s.PCT, s.SC + s.CT + s.PS + s.PCT AS TotalHoras, m.Descricao, DATEPART(WEEK, CAST(s.Data AS date)) as Semana, a.Codigo_Curso, m.Nivel FROM TBForSessoes s INNER JOIN TBForAccoes a ON s.Rowid_Accao = a.versao_rowid INNER JOIN TBForModulos m ON s.Rowid_Modulo = m.versao_rowid WHERE (a.versao_rowid = '" + rowidAccao + "') order by LEN(s.Num_Sessao), s.Num_Sessao", connection);
            DataTable dataTable = new DataTable();
            sqlDataAdapter.Fill(dataTable);
            List<DataRow> list = dataTable.AsEnumerable().ToList<DataRow>();
            connection.Close();
            return list;
        }
        public static List<DataRow> getModulosHT(string codCurso, string refCurso, SqlConnection connection)
        {
            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter("SELECT m.versao_rowid, m.Nivel, m.Descricao, m.N_Horas, m.Peso_Aval, s.Data, s.Hora_Inicio, s.Hora_Fim, ROW_NUMBER() OVER(ORDER BY s.Hora_Fim DESC) as RowNum FROM TBForModulos m INNER JOIN TBForSessoes s ON m.versao_rowid = s.Rowid_Modulo INNER JOIN TBForAccoes a ON s.Rowid_Accao = a.versao_rowid WHERE(m.Codigo_Curso = '" + codCurso + "') AND(a.Ref_Accao = '" + refCurso + "')", connection);
            DataTable dataTable = new DataTable();
            sqlDataAdapter.Fill(dataTable);
            List<DataRow> list = dataTable.AsEnumerable().ToList<DataRow>();
            connection.Close();
            return list;
        }
        public static string getFormandosBloqueados(string datainicio, string datafim, string codformador)
        {
            return "select f.Nome_Abreviado, i.Data_Desistencia from TBForInscricoes i JOIN TBForFormandos f ON (i.Codigo_Formando = f.Codigo_Formando ) JOIN TBForSessoes s ON(s.Codigo_Formador = " + codformador + " and Data BETWEEN '" + datainicio + "' and '" + datafim + "') where i.Desistencia_motivo = 14 and i.Rowid_Accao = s.Rowid_Accao"; 
        }
        public static List<DataRow> getSelectedTurma(string refAccao, SqlConnection connection)
        {
            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter("SELECT Formandos.Codigo_Formando, Formandos.Formando, Formandos.Data_Nascimento, Formandos.BI, Formandos.BI_Data_Validade, Formandos.NC, Moradas.Codigo_Pais, Moradas.Rua, Moradas.Localidade, Moradas.Codigo_Postal, Moradas.SubCodigo_Postal, Moradas.Descricao_Postal, Contactos.Telefone1, Contactos.Telefone2, Contactos.Email1, Contactos.Email2, Formandos.Sexo, Formandos.Codigo_Formando_2 FROM (SELECT versao_rowid, Codigo_Formando, Formando, Data_Nascimento, BI, BI_Data_Validade, NC, Sexo, Codigo_Formando_2 FROM TBForFormandos) as Formandos, (SELECT Codigo_Entidade, Codigo_Pais, Rua, Localidade, Codigo_Postal, SubCodigo_Postal, Descricao_Postal FROM TBGerMoradas where Tipo_Entidade = 3) as Moradas, (SELECT Codigo_Entidade, Telefone1, Telefone2, Email1, Email2 FROM TBGerContactos  where Tipo_Entidade = 3) as Contactos, (SELECT a.codigo_curso, f.codigo_formando, f.formando, a.Ref_Accao FROM TBForInscricoes i INNER JOIN TBForFormandos f ON i.codigo_formando = f.codigo_formando and i.confirmado = 1 and i.desistente = 0 INNER JOIN TBForAccoes a ON a.versao_rowid = i.rowid_accao INNER JOIN TBForCursos c ON a.codigo_curso = c.codigo_curso WHERE a.Ref_Accao = '" + refAccao + "') as Turmas where Formandos.versao_rowid = Moradas.Codigo_Entidade AND Formandos.versao_rowid = Contactos.Codigo_Entidade AND Turmas.Codigo_Formando = Formandos.Codigo_Formando order by Formandos.Formando", connection);
            DataTable dataTable = new DataTable();
            sqlDataAdapter.Fill(dataTable);
            List<DataRow> list = dataTable.AsEnumerable().ToList<DataRow>();
            connection.Close();
            return list;
        }
        public static List<DataRow> getFormadores(string refAccao, SqlConnection connection)
        {
            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter("SELECT DISTINCT(TBForFormadores.Codigo_Formador), TBForFormadores.Nome_Abreviado, TBGerContactos.Email1, TBGerContactos.Telefone1 " + "FROM TBForSessoes INNER JOIN " + "TBForAccoes ON TBForSessoes.Rowid_Accao = TBForAccoes.versao_rowid INNER JOIN " + "TBForFormadores ON TBForSessoes.Codigo_Formador = TBForFormadores.Codigo_Formador INNER JOIN " + "TBGerContactos ON TBForFormadores.versao_rowid = TBGerContactos.Codigo_Entidade INNER JOIN " + "TBForCandAccoes ON TBForAccoes.Numero_Projecto = TBForCandAccoes.Numero_Cand AND TBForAccoes.Codigo_Curso = TBForCandAccoes.Codigo_Curso AND " + "TBForAccoes.Numero_Accao = TBForCandAccoes.Numero_Accao " + "WHERE(TBForAccoes.Ref_Accao = '" + refAccao + "') AND(TBGerContactos.Tipo_Entidade = 4)  AND TBForFormadores.Codigo_Formador NOT IN (699, 704, 827, 1046) " + "group by TBForFormadores.Codigo_Formador, TBForFormadores.Nome_Abreviado, TBGerContactos.Email1, TBGerContactos.Telefone1", connection);
            DataTable dataTable = new DataTable();
            sqlDataAdapter.Fill(dataTable);
            List<DataRow> list = dataTable.AsEnumerable().ToList<DataRow>();
            connection.Close();
            return list;
        }
        public static List<DataRow> getFormadoresAll(SqlConnection connection)
        {
            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter("SELECT DISTINCT(TBForFormadores.Codigo_Formador), TBForFormadores.Nome_Abreviado, TBGerContactos.Email1, TBGerContactos.Email2, TBGerContactos.Telefone1, TBForFormadores.Sexo " + "FROM TBForSessoes INNER JOIN " + "TBForAccoes ON TBForSessoes.Rowid_Accao = TBForAccoes.versao_rowid INNER JOIN " + "TBForFormadores ON TBForSessoes.Codigo_Formador = TBForFormadores.Codigo_Formador INNER JOIN " + "TBGerContactos ON TBForFormadores.versao_rowid = TBGerContactos.Codigo_Entidade INNER JOIN " + "TBForCandAccoes ON TBForAccoes.Numero_Projecto = TBForCandAccoes.Numero_Cand AND TBForAccoes.Codigo_Curso = TBForCandAccoes.Codigo_Curso AND " + "TBForAccoes.Numero_Accao = TBForCandAccoes.Numero_Accao " + "WHERE (TBGerContactos.Tipo_Entidade = 4)  AND TBForFormadores.Codigo_Formador NOT IN (699, 704, 827, 1046) " + "group by TBForFormadores.Codigo_Formador, TBForFormadores.Nome_Abreviado, TBGerContactos.Email1,TBGerContactos.Email2, TBGerContactos.Telefone1, TBForFormadores.Sexo", connection);
            DataTable dataTable = new DataTable();
            sqlDataAdapter.Fill(dataTable);
            List<DataRow> list = dataTable.AsEnumerable().ToList<DataRow>();
            connection.Close();
            return list;
        }
        public static List<DataRow> getFormadores_Agenda(string refAccao, SqlConnection connection)
        {
            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter("SELECT DISTINCT(TBForFormadores.Codigo_Formador), TBForFormadores.Nome_Abreviado, TBGerContactos.Email1, TBGerContactos.Email2, TBGerContactos.Telefone1 " + "FROM TBForSessoes INNER JOIN " + "TBForAccoes ON TBForSessoes.Rowid_Accao = TBForAccoes.versao_rowid INNER JOIN " + "TBForFormadores ON TBForSessoes.Codigo_Formador = TBForFormadores.Codigo_Formador INNER JOIN " + "TBGerContactos ON TBForFormadores.versao_rowid = TBGerContactos.Codigo_Entidade INNER JOIN " + "TBForCandAccoes ON TBForAccoes.Numero_Projecto = TBForCandAccoes.Numero_Cand AND TBForAccoes.Codigo_Curso = TBForCandAccoes.Codigo_Curso AND " + "TBForAccoes.Numero_Accao = TBForCandAccoes.Numero_Accao " + "WHERE(TBForAccoes.Ref_Accao = '" + refAccao + "') AND(TBGerContactos.Tipo_Entidade = 4)  AND TBForFormadores.Codigo_Formador NOT IN (699, 704, 827, 1046) " + "group by TBForFormadores.Codigo_Formador, TBForFormadores.Nome_Abreviado, TBGerContactos.Email1, TBGerContactos.Email2, TBGerContactos.Telefone1", connection);
            DataTable dataTable = new DataTable();
            sqlDataAdapter.Fill(dataTable);
            List<DataRow> list = dataTable.AsEnumerable().ToList<DataRow>();
            connection.Close();
            return list;
        }
        public static string getFormadorestelefone(string datainicio, string datafim)
        {
            return "SELECT DISTINCT(TBForFormadores.Codigo_Formador) as Codigo_Formador, TBGerContactos.Telefone1, TBForFormandos.Formando, TBForInscricoes.Data_Desistencia, TBForInscricoes.Desistencia_motivo FROM TBForSessoes INNER JOIN TBForAccoes ON TBForSessoes.Rowid_Accao = TBForAccoes.versao_rowid and TBForSessoes.Data BETWEEN '" + datainicio + "' and '" + datafim + "' INNER JOIN TBForInscricoes ON TBForInscricoes.Rowid_Accao=TBForSessoes.Rowid_Accao INNER JOIN TBForFormadores ON TBForSessoes.Codigo_Formador = TBForFormadores.Codigo_Formador INNER JOIN TBForFormandos ON (TBForInscricoes.Codigo_Formando = tbforformandos.Codigo_Formando )  INNER JOIN TBGerContactos ON TBForFormadores.versao_rowid = TBGerContactos.Codigo_Entidade WHERE TBGerContactos.Tipo_Entidade = 4  AND TBForFormadores.Codigo_Formador NOT IN (699, 704, 827, 1046) group by TBForFormadores.Codigo_Formador, TBForFormadores.Formador, TBGerContactos.Email1, TBGerContactos.Telefone1, TBForFormandos.Formando, TBForInscricoes.Data_Desistencia, TBForInscricoes.Desistencia_motivo"; 
        }
        public static List<DataRow> getModulosHT2(string refAccao, SqlConnection connection)
        {
            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter("SELECT TBForModulos.versao_rowid, TBForSessoes.Hora_Inicio, TBForModulos.Descricao, TBForModulos.N_Horas, TBForFormadores.Nome_Abreviado, TBGerContactos.Email1, TBGerContactos.Telefone1, " + "TBForFormadores.Codigo_Formador, TBForLocais.Local, TBForCursos.Descricao, TBForAccoes.Ref_Accao, TBForCursos.Conteudos_Prog, TBForModulos.Nivel FROM TBForModulos INNER JOIN " + "TBForSessoes ON TBForModulos.versao_rowid = TBForSessoes.Rowid_Modulo INNER JOIN " + "TBForAccoes ON TBForSessoes.Rowid_Accao = TBForAccoes.versao_rowid INNER JOIN " + "TBForFormadores ON TBForSessoes.Codigo_Formador = TBForFormadores.Codigo_Formador INNER JOIN " + "TBGerContactos ON TBForFormadores.versao_rowid = TBGerContactos.Codigo_Entidade INNER JOIN " + "TBForCandAccoes ON TBForAccoes.Numero_Projecto = TBForCandAccoes.Numero_Cand AND TBForAccoes.Codigo_Curso = TBForCandAccoes.Codigo_Curso AND " + "TBForAccoes.Numero_Accao = TBForCandAccoes.Numero_Accao INNER JOIN " + "TBForLocais ON TBForCandAccoes.Codigo_Local = TBForLocais.Codigo_Local INNER JOIN TBForCursos ON TBForAccoes.Codigo_Curso=TBForCursos.Codigo_Curso " + "WHERE(TBForAccoes.Ref_Accao = '" + refAccao + "') AND (TBGerContactos.Tipo_Entidade = 4) " + "ORDER BY TBForSessoes.Hora_Inicio ", connection);
            DataTable dataTable = new DataTable();
            sqlDataAdapter.Fill(dataTable);
            List<DataRow> list = dataTable.AsEnumerable().ToList<DataRow>();
            connection.Close();
            return list;
        }
        public static List<DataRow> getModulosHT_Formador(string FormadorID, SqlConnection connection)
        {
            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter("SELECT TBForModulos.versao_rowid, TBForSessoes.Hora_Inicio, TBForModulos.Descricao, TBForModulos.N_Horas, TBForFormadores.Nome_Abreviado, TBGerContactos.Email1, TBGerContactos.Telefone1, TBForFormadores.Codigo_Formador, TBForLocais.Local, TBForCursos.Descricao, TBForAccoes.Ref_Accao, TBForCursos.Conteudos_Prog, TBForModulos.Nivel FROM TBForModulos INNER JOIN TBForSessoes ON TBForModulos.versao_rowid = TBForSessoes.Rowid_Modulo INNER JOIN TBForAccoes ON TBForSessoes.Rowid_Accao = TBForAccoes.versao_rowid INNER JOIN TBForFormadores ON TBForSessoes.Codigo_Formador = TBForFormadores.Codigo_Formador INNER JOIN TBGerContactos ON TBForFormadores.versao_rowid = TBGerContactos.Codigo_Entidade INNER JOIN TBForCandAccoes ON TBForAccoes.Numero_Projecto = TBForCandAccoes.Numero_Cand AND TBForAccoes.Codigo_Curso = TBForCandAccoes.Codigo_Curso AND TBForAccoes.Numero_Accao = TBForCandAccoes.Numero_Accao INNER JOIN TBForLocais ON TBForCandAccoes.Codigo_Local = TBForLocais.Codigo_Local INNER JOIN TBForCursos ON TBForAccoes.Codigo_Curso=TBForCursos.Codigo_Curso WHERE(TBForFormadores.Codigo_Formador = '" + FormadorID + "') AND (TBGerContactos.Tipo_Entidade = 4) " + "ORDER BY TBForSessoes.Hora_Inicio ", connection);
            DataTable dataTable = new DataTable();
            sqlDataAdapter.Fill(dataTable);
            List<DataRow> list = dataTable.AsEnumerable().ToList<DataRow>();
            connection.Close();
            return list;
        }
    }
}