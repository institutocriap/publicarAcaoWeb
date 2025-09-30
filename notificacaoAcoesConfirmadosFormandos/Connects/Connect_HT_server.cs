using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NotificacaoAcoesConfirmadasFormandos.Connects
{
    class Connect_HT_server
    {
        private string DBhost;
        private string DBname;
        private string DBuser;
        private string DBpass;
        private SqlConnection Connection;
        public static string ConnectionString = null;

        public Connect_HT_server(string host, string dbname, string user, string pass)
        {
            DBhost = host;
            DBname = dbname;
            DBuser = user;
            DBpass = pass;

            ConnectionString = "Data Source='" + DBhost + "'; Initial Catalog='" + DBname + "'; User Id='" + DBuser + "'; Password='" + DBpass + "'; Trusted_Connection=False";

            Connection = new SqlConnection();
            Connection.ConnectionString = ConnectionString;
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
            Connection.Close();
            Connection.Open();
        }
        public void ConnEnd()
        {
            Connection.Close();
        }
        public SqlConnection Conn { get { return Connection; } }
    }
}