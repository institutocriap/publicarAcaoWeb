using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace NotificacaoAcoesConfirmadasFormandos.Connects
{
    class Security
    {
        private static readonly byte[] SALT = new byte[] { 0x26, 0xdc, 0xff, 0x00, 0xad, 0xed, 0x7a, 0xee, 0xc5, 0xfe, 0x07, 0xaf, 0x4d, 0x08, 0x22, 0x3c };
        public static readonly string KEY = "ioudfndsoiw3jk82bdnbiuncdsilolkj";

        public static appSettings settings = new appSettings();

        public static string EncryptText(string toEncrypt, bool useHashing)
        {
            byte[] keyArray;
            byte[] toEncryptArray = UTF8Encoding.UTF8.GetBytes(toEncrypt);

            string key = "jgklfjd9!8u5nf9098545jk34";
            if (useHashing)
            {
                MD5CryptoServiceProvider hashmd5 = new MD5CryptoServiceProvider();
                keyArray = hashmd5.ComputeHash(UTF8Encoding.UTF8.GetBytes(key));
                hashmd5.Clear();
            }
            else
                keyArray = UTF8Encoding.UTF8.GetBytes(key);

            TripleDESCryptoServiceProvider tdes = new TripleDESCryptoServiceProvider();
            tdes.Key = keyArray;
            tdes.Mode = CipherMode.ECB;
            tdes.Padding = PaddingMode.PKCS7;

            ICryptoTransform cTransform = tdes.CreateEncryptor();
            byte[] resultArray =
              cTransform.TransformFinalBlock(toEncryptArray, 0,
              toEncryptArray.Length);
            tdes.Clear();
            return Convert.ToBase64String(resultArray, 0, resultArray.Length);
        }

        public static string DecryptText(string cipherString, bool useHashing)
        {
            byte[] keyArray;
            byte[] toEncryptArray = Convert.FromBase64String(cipherString);

            string key = "jgklfjd9!8u5nf9098545jk34";

            if (useHashing)
            {
                MD5CryptoServiceProvider hashmd5 = new MD5CryptoServiceProvider();
                keyArray = hashmd5.ComputeHash(UTF8Encoding.UTF8.GetBytes(key));

                hashmd5.Clear();
            }
            else
            {
                keyArray = UTF8Encoding.UTF8.GetBytes(key);
            }

            TripleDESCryptoServiceProvider tdes = new TripleDESCryptoServiceProvider();
            tdes.Key = keyArray;

            tdes.Mode = CipherMode.ECB;
            tdes.Padding = PaddingMode.PKCS7;

            ICryptoTransform cTransform = tdes.CreateDecryptor();
            byte[] resultArray = cTransform.TransformFinalBlock(
                                 toEncryptArray, 0, toEncryptArray.Length);

            tdes.Clear();
            return UTF8Encoding.UTF8.GetString(resultArray);
        }

        public static void localhost()
        {
            geral();
            settings.wp_HName = EncryptText("localhost", true);
            settings.wp_DBName = EncryptText("site_2017", true);
            settings.wp_UName = EncryptText("root", true);
            settings.wp_Pass = EncryptText("fgdfgfdg", true);
        }

        public static void remote()
        {
            geral();
        }

        private static void geral()
        {
            string path = @"\\192.168.1.248\docs\SV";
            string arquivo = Path.Combine(path, "settings.ini");

            string[] lines = System.IO.File.ReadAllLines(@"\\192.168.1.248\docs\SV\settings.ini");

            int lineposicao = 0;
            foreach (string line in lines)
            {
                string[] words = line.Split('|');
                if (lineposicao == 8) { settings.ht_HName = DecryptText(words[1].ToString(), true); }
                if (lineposicao == 9) { settings.ht_DBName = DecryptText(words[1].ToString(), true); }
                if (lineposicao == 10) { settings.ht_UName = DecryptText(words[1].ToString(), true); }
                if (lineposicao == 11) { settings.ht_Pass = DecryptText(words[1].ToString(), true); }

                if (lineposicao == 26) { Properties.Settings.Default["emailenvio"] = DecryptText(words[1].ToString(), true); }
                if (lineposicao == 27) { Properties.Settings.Default["passwordemail"] = DecryptText(words[1].ToString(), true); }
                lineposicao++;
            }
            Properties.Settings.Default["ConnectionString"] = "Data Source='" + settings.ht_HName + "'; Initial Catalog='" + settings.ht_DBName + "'; User Id=" + settings.ht_UName + "; Password=" + settings.ht_Pass + "; ";
            Properties.Settings.Default.Save();
        }
        public static byte[] EncryptFile(byte[] clearData, string password)
        {
            MemoryStream ms = new MemoryStream();
            Rijndael alg = Rijndael.Create();
            Rfc2898DeriveBytes pdb = new Rfc2898DeriveBytes(password, SALT);

            alg.Key = pdb.GetBytes(32);
            alg.IV = pdb.GetBytes(16);

            CryptoStream cs = new CryptoStream(ms, alg.CreateEncryptor(), CryptoStreamMode.Write);
            cs.Write(clearData, 0, clearData.Length);
            cs.Close();
            byte[] encryptedData = ms.ToArray();
            return encryptedData;
        }
        public static byte[] DecryptFile(byte[] cipherData, string password)
        {
            MemoryStream ms = new MemoryStream();
            Rijndael alg = Rijndael.Create();
            Rfc2898DeriveBytes pdb = new Rfc2898DeriveBytes(password, SALT);

            alg.Key = pdb.GetBytes(32);
            alg.IV = pdb.GetBytes(16);

            CryptoStream cs = new CryptoStream(ms, alg.CreateDecryptor(), CryptoStreamMode.Write);
            cs.Write(cipherData, 0, cipherData.Length);
            cs.Close();
            byte[] decryptedData = ms.ToArray();
            return decryptedData;
        }
        public class appSettings
        {
            public string pa_HName { get; set; }
            public string pa_DBName { get; set; }
            public string pa_UName { get; set; }
            public string pa_Pass { get; set; }
            public string ht_HName { get; set; }
            public string ht_DBName { get; set; }
            public string ht_UName { get; set; }
            public string ht_Pass { get; set; }
            public string wp_HName { get; set; }
            public string wp_DBName { get; set; }
            public string wp_UName { get; set; }
            public string wp_Pass { get; set; }
        }
    }
}