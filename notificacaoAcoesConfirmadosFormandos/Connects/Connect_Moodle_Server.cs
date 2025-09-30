using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MySql.Data.MySqlClient;

namespace SalasZoomNotificationFormadores.connects
{
    class Connect_Moodle_Server
    {
        private string DBserver;
        private string DBuid;
        private string DBpwd;
        private string DBdatabase;
        private MySqlConnection Connection;
        public static string ConnectionString = null;

        public Connect_Moodle_Server(string server, string uid, string pwd, string database)
        {
            DBserver = server;
            DBuid = uid;
            DBpwd = pwd;
            DBdatabase = database;

            ConnectionString = "server='" + server + "';uid='" + uid + "';pwd='" + pwd + "';database='" + database + "'; pooling = false; convert zero datetime=True; default command timeout=20;";

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
            Connection.Close();
            Connection.Open();
        }

        public void ConnEnd()
        {
            Connection.Close();
        }

        public MySqlConnection Conn { get { return Connection; } }
    }
}

