using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using log4net;
//using System.Data.OracleClient;
using Oracle.DataAccess.Client;

namespace SEI
{
    class Settings
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        
        List<SEIUser> sEIUsers = null;
        public List<SEIUser> SEIUsers { get { return sEIUsers; } }

        public List<String> Rooms;

        public Settings()
        {
            try
            {
                OracleConnectionStringBuilder oCSB = new OracleConnectionStringBuilder
                {
                    UserID = Properties.Settings.Default.OracleLogin,
                    DataSource = Properties.Settings.Default.OracleTNS,
                    Password = Properties.Settings.Default.uOraclePassword,
                    Pooling = true,
                    ConnectionTimeout = Properties.Settings.Default.ConnectionTimeout, //secs wait for new connection
                    ConnectionLifeTime = Properties.Settings.Default.ConnectionLifeTime, //secs reuse or not
                    DecrPoolSize = Properties.Settings.Default.DecrPoolSize,
                    IncrPoolSize = Properties.Settings.Default.IncrPoolSize,
                    MinPoolSize = Properties.Settings.Default.MinPoolSize,
                    MaxPoolSize = Properties.Settings.Default.MaxPoolSize
                };

                Properties.Settings.Default.uOracleConnectionString = oCSB.ConnectionString;


                oCSB.ConnectionTimeout = Properties.Settings.Default.ConnectionTimeoutS;
                oCSB.ConnectionLifeTime = Properties.Settings.Default.ConnectionLifeTimeS;
                oCSB.DecrPoolSize = Properties.Settings.Default.DecrPoolSizeS;
                oCSB.IncrPoolSize = Properties.Settings.Default.IncrPoolSizeS;
                oCSB.MinPoolSize = Properties.Settings.Default.MinPoolSizeS;
                oCSB.MaxPoolSize = Properties.Settings.Default.MaxPoolSizeS;
                Properties.Settings.Default.uOracleConnectionStringS = oCSB.ConnectionString;
            }
            catch (Exception e)
            {
                if (log.IsErrorEnabled) log.Error("OracleConnectionStringBuilder: " + e.Message);

            }
        }

        public void ReloadSettings()
        {
            using (ThreadContext.Stacks["NDC"].Push("Reloading settings"))
            {

                OracleConnection oConnection = new OracleConnection();
                OracleCommand oCommand;
                OracleDataReader oReader;
                if (sEIUsers == null)
                    sEIUsers = new List<SEIUser>();
                else
                    sEIUsers.Clear();
                try
                {
                    oConnection = new OracleConnection(Properties.Settings.Default.uOracleConnectionStringS);
                    oConnection.Open();
                    oCommand = new OracleCommand
                    {
                        BindByName = true,
                        Connection = oConnection,
                        CommandType = System.Data.CommandType.Text,
                        //oCommand.CommandText = "select user_login, email_addr, direction_id, sync_start_date,SYNC_TOKEN_l, LAST_SYNC_DATE, USER_ID from " + Properties.Settings.Default.OracleSchema + ".cx_sei_users";
                        //My
                        CommandText = "select user_login, email_addr, direction_id, sync_start_date,SYNC_TOKEN, LAST_SYNC_DATE, USER_ID from " + Properties.Settings.Default.OracleSchema + ".cx_sei_users"
                    };

                    oReader = oCommand.ExecuteReader();
                    while (oReader.Read())
                    {
                        sEIUsers.Add(new SEIUser(oReader.GetString(0), oReader.GetString(1), oReader.GetString(2), oReader.GetDateTime(3).ToLocalTime(), oReader.IsDBNull(4) ? "" : oReader.GetString(4), oReader.IsDBNull(5) ? null : (new DateTime?(oReader.GetDateTime(5).ToLocalTime())), oReader.GetString(6)));
                    }
                    if (log.IsInfoEnabled) log.Info("Settings loaded");
                    //Logger.Log("Settings loaded");

                    oConnection.Close();
                    oConnection.Dispose();
                }
                catch (Exception e)
                {
                    log.Error("Error loading settings: " + e.Message);
                    //Logger.Log("Error loading settings: " + e.Message, Logger.LogSeverity.Error);
                }
                finally
                {
                    oConnection.Close();
                    oConnection.Dispose();
                }
            }
        }
    }
}
