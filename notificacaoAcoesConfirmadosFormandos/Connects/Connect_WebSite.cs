using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NotificacaoAcoesConfirmadasFormandos.Connects
{
    class Connect_WebSite
    {
        private string DBserver;
        private string DBuid;
        private string DBpwd;
        private string DBdatabase;
        private MySqlConnection Connection;
        public static string ConnectionString = null;

        public Connect_WebSite(string server, string uid, string pwd, string database)
        {
            DBserver = server;
            DBuid = uid;
            DBpwd = pwd;
            DBdatabase = database;
            ConnectionString = "server=" + server + ";uid=" + uid + ";pwd=" + pwd + ";database=" + database + "; default command timeout=20;Convert Zero Datetime=True";
            Connection = new MySqlConnection(ConnectionString);
        }

        public Boolean TestConnect()
        {
            try
            {
                Connection.Open();
                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
                Connection.Close();
            }
        }

        public void ConnInit()
        {
            try
            {
                Connection.Close();
                Connection.Open();
            }
            catch (Exception e)
            {
                MessageBox.Show("Ocorreu um erro WebSite Server!\n\n" + e.ToString(), (string)Properties.Settings.Default["title"], MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        public void ConnEnd()
        {
            Connection.Close();
        }

        public MySqlConnection Conn { get { return Connection; } }
    }
}
