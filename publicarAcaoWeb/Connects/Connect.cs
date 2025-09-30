using SalasZoomNotificationFormadores.connects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PublicarAcaoWeb.Connects
{
    class Connect
    {
        public static Connect_HT_server HTlocalConnect = new Connect_HT_server(Security.settings.ht_HName, Security.settings.ht_DBName, Security.settings.ht_UName, Security.settings.ht_Pass);
        public static Connect_HT_server SVlocalConnect = new Connect_HT_server(Security.settings.ht_HName, "secretariaVirtual", Security.settings.ht_UName, Security.settings.ht_Pass);
        public static Connect_Moodle_Server MoodleConnect = new Connect_Moodle_Server("94.46.28.131", "eadcriap_eadmood", "ZWBd*F;xn!yX", "eadcriap_eadmood");
        public static Connect_WebSite webSiteConnect = new Connect_WebSite("94.46.28.119", "sucriap_dev", "nyoWq+h~ha~4", "sucriap_suitecrm");

        public static void closeAll()
        {
            HTlocalConnect.Conn.Close();
            SVlocalConnect.Conn.Close();
            MoodleConnect.Conn.Close();
            webSiteConnect.Conn.Close();
        }
    }
}
