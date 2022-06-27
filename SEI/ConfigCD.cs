using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Security.Cryptography;
using System.IO;
using SEI;
using log4net;

namespace SEI
{
    class ConfigCD
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        static readonly string PasswordHash = "xKUg78HJGF_787&^%892kf";
        static readonly string SaltKey = "FKSGHDP#$@*GVSWO39s2";
        static readonly string VIKey = "j&gx2LS9D($e3p9d";

        private static string Encrypt(string plainText)
        {           
            //Logger.Log("Encrypting...");
            byte[] plainTextBytes = Encoding.UTF8.GetBytes(plainText);

            byte[] keyBytes = new Rfc2898DeriveBytes(PasswordHash, Encoding.ASCII.GetBytes(SaltKey)).GetBytes(256 / 8);
            var symmetricKey = new RijndaelManaged() { Mode = CipherMode.CBC, Padding = PaddingMode.Zeros };
            var encryptor = symmetricKey.CreateEncryptor(keyBytes, Encoding.ASCII.GetBytes(VIKey));

            byte[] cipherTextBytes;

            using (var memoryStream = new MemoryStream())
            {
                using (var cryptoStream = new CryptoStream(memoryStream, encryptor, CryptoStreamMode.Write))
                {
                    cryptoStream.Write(plainTextBytes, 0, plainTextBytes.Length);
                    cryptoStream.FlushFinalBlock();
                    cipherTextBytes = memoryStream.ToArray();
                    cryptoStream.Close();
                }
                memoryStream.Close();
            }
            return Convert.ToBase64String(cipherTextBytes);
        }
        public static string Decrypt(string encryptedText)
        {
            //Logger.Log("Decrypting...");
            byte[] cipherTextBytes = Convert.FromBase64String(encryptedText);
            byte[] keyBytes = new Rfc2898DeriveBytes(PasswordHash, Encoding.ASCII.GetBytes(SaltKey)).GetBytes(256 / 8);
            var symmetricKey = new RijndaelManaged() { Mode = CipherMode.CBC, Padding = PaddingMode.None };

            var decryptor = symmetricKey.CreateDecryptor(keyBytes, Encoding.ASCII.GetBytes(VIKey));
            var memoryStream = new MemoryStream(cipherTextBytes);
            var cryptoStream = new CryptoStream(memoryStream, decryptor, CryptoStreamMode.Read);
            byte[] plainTextBytes = new byte[cipherTextBytes.Length];

            int decryptedByteCount = cryptoStream.Read(plainTextBytes, 0, plainTextBytes.Length);
            memoryStream.Close();
            cryptoStream.Close();
            return Encoding.UTF8.GetString(plainTextBytes, 0, decryptedByteCount).TrimEnd("\0".ToCharArray());
        }

        public static bool TryDecrypt(string encryptedText)
        {
            bool ret = true;
            try
            {
                byte[] cipherTextBytes = Convert.FromBase64String(encryptedText);
                byte[] keyBytes = new Rfc2898DeriveBytes(PasswordHash, Encoding.ASCII.GetBytes(SaltKey)).GetBytes(256 / 8);
                var symmetricKey = new RijndaelManaged() { Mode = CipherMode.CBC, Padding = PaddingMode.None };

                var decryptor = symmetricKey.CreateDecryptor(keyBytes, Encoding.ASCII.GetBytes(VIKey));
                var memoryStream = new MemoryStream(cipherTextBytes);
                var cryptoStream = new CryptoStream(memoryStream, decryptor, CryptoStreamMode.Read);
                byte[] plainTextBytes = new byte[cipherTextBytes.Length];

                int decryptedByteCount = cryptoStream.Read(plainTextBytes, 0, plainTextBytes.Length);
                memoryStream.Close();
                cryptoStream.Close();
                Encoding.UTF8.GetString(plainTextBytes, 0, decryptedByteCount).TrimEnd("\0".ToCharArray());
            }
#pragma warning disable CS0168 // Переменная объявлена, но не используется
            catch (Exception e)
            {
                ret = false;
            }
            return ret;
        }

        public static void CheckConfig()
        {
            using (ThreadContext.Stacks["NDC"].Push("Checking configuration"))
            {

                System.Configuration.Configuration configFile = System.Configuration.ConfigurationManager.OpenExeConfiguration(System.Configuration.ConfigurationUserLevel.None);
                System.Configuration.ClientSettingsSection section = (System.Configuration.ClientSettingsSection)configFile.SectionGroups["applicationSettings"].Sections[0];
                log.Info("Configuration Loaded");

                if (!TryDecrypt(Properties.Settings.Default.SiebelPassword))
                {
                    log.Warn("Siebel password unencrypted");
                    section.Settings.Get("SiebelPassword").Value.ValueXml.InnerXml = Encrypt(Properties.Settings.Default.SiebelPassword);
                }
                if (!TryDecrypt(Properties.Settings.Default.OraclePassword))
                {
                    log.Warn("Oracle password unencrypted");
                    section.Settings.Get("OraclePassword").Value.ValueXml.InnerXml = Encrypt(Properties.Settings.Default.OraclePassword);
                }
                if (!TryDecrypt(Properties.Settings.Default.ExchangePassword))
                {
                    log.Warn("Exchange password unencrypted");
                    section.Settings.Get("ExchangePassword").Value.ValueXml.InnerXml = Encrypt(Properties.Settings.Default.ExchangePassword);
                }

                section.SectionInformation.ForceSave = true;
                configFile.Save(System.Configuration.ConfigurationSaveMode.Full);
                //
                //if (log.IsDebugEnabled) log.Debug("Config file saved");
                //


                Properties.Settings.Default.Reload();

                Properties.Settings.Default.uSiebelPassword = Decrypt(Properties.Settings.Default.SiebelPassword);

                //
                //if (log.IsDebugEnabled) log.Debug( "Siebel password: " + Properties.Settings.Default.uSiebelPassword);
                //

                Properties.Settings.Default.uOraclePassword = Decrypt(Properties.Settings.Default.OraclePassword);

                //
                //if (log.IsDebugEnabled) log.Debug("Oracle password: " + Properties.Settings.Default.uOraclePassword);
                //

                Properties.Settings.Default.uExchangePassword = Decrypt(Properties.Settings.Default.ExchangePassword);

                //
                //if (log.IsDebugEnabled) log.Debug("Exchange password: " + Properties.Settings.Default.uExchangePassword);
                //

            }
        }
    }
}
