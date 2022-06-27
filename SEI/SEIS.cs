using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using log4net.Layout;
using log4net;


namespace SEI
{
    public partial class SEIS : ServiceBase
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private Settings allSettings;
        private System.Timers.Timer timer;
        object timerLock = new Object();
        public LogApache lgAppache;
        private bool IsSEIUsersMustUpdate;

        public SEIS()
        {
            lgAppache = new LogApache();
            IsSEIUsersMustUpdate = false;
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            using (ThreadContext.Stacks["NDC"].Push("Service starting"))
            {
                if (log.IsInfoEnabled) log.Info("Service Started");
                ConfigCD.CheckConfig();
                
                //
                //if (log.IsDebugEnabled) log.Debug("CheckConfig passed");
                //
                try
                {
                    allSettings = new Settings();
                }

                catch (Exception e)
                {
                    if (log.IsErrorEnabled) log.Error("Settings: " + e.Message);                  

                }

                this.timer = new System.Timers.Timer((Double)Properties.Settings.Default.SyncPeriod * 1000);
                this.timer.AutoReset = true;
                this.timer.Elapsed += new System.Timers.ElapsedEventHandler(this.timer_Elapsed);

                //
                //if (log.IsDebugEnabled) log.Debug("Starting timer ...");
                //
                try
                {
                    this.timer.Start();
                }
                catch (Exception e)
                {
                    if (log.IsErrorEnabled) log.Error("Catch: " + e.Message);
                }
                //
                //if (log.IsDebugEnabled) log.Debug("OnStart completed !!!");
                //

            }
        }

        private void timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            using (ThreadContext.Stacks["NDC"].Push("Synchronization cycle"))
            {
                ////
                //if (log.IsDebugEnabled) log.Debug("In timer elapsed !!!");
                ////

                if (!System.Threading.Monitor.TryEnter(timerLock))
                    return;
                if (!String.IsNullOrEmpty(Properties.Settings.Default.TimeSEIUsersSync))
                {
                    try
                    {
                        TimeSpan propSyncTime = new TimeSpan(System.Convert.ToInt16(Properties.Settings.Default.TimeSEIUsersSync.Substring(0, 2)),
                            System.Convert.ToInt16(Properties.Settings.Default.TimeSEIUsersSync.Substring(3, 2)), 0);
                        if (e.SignalTime.TimeOfDay <= propSyncTime) IsSEIUsersMustUpdate = true;
                        if (e.SignalTime.TimeOfDay >= propSyncTime && IsSEIUsersMustUpdate)
                        {
                            if (log.IsInfoEnabled) log.Info("------------------------------------------------------------------");
                            if (log.IsInfoEnabled) log.Info(" List of SEI users updating");
                            IsSEIUsersMustUpdate = false;
                            SiebelDBW.UpdateListUsers();

                        }
                    }
                    catch (Exception ex)
                    {
                        log.Error("Error SEI users updating: " + ex.Message);
                    }
                }
                //SiebelDBW.UpdateListUsers();
                List<Task> syncTasks = new List<Task>();

                bool bAllOK = true;
                try
                {
                    if (log.IsInfoEnabled) log.Info("------------------------------------------------------------------");
                    if (log.IsInfoEnabled) log.Info("Synchronization start");

                    allSettings.ReloadSettings();
                    if (allSettings.SEIUsers.Count > 0)
                    {
                        bAllOK = Syncee.LoadRooms(out allSettings.Rooms);
                        if (bAllOK)
                            bAllOK = Syncee.SyncSiebelDeleted();
                    }

                    List<string> conUsers = new List<string>();
                    List<string> conUpperUsersMail = new List<string>();
                    //if (log.IsDebugEnabled) log.Debug("Getting max last sync time");
                    DateTime maxDTLastSync =  SiebelDBW.GetMaxDTLastSync();                    
                    //if (log.IsDebugEnabled) log.Debug("Max last sync time - " + maxDTLastSync);
                    if (maxDTLastSync != DateTime.MaxValue)
                    {
                        conUsers = SiebelDBW.GetConsiderUserList(maxDTLastSync - new TimeSpan(1,0,0));
                    }                                    
                    foreach (string email in conUsers)
                    {
                        conUpperUsersMail.Add(email.ToUpper());
                    }
                    if (log.IsDebugEnabled && conUpperUsersMail.Count > 0)
                    {
                        foreach (string email in conUsers)
                        {
                            log.Debug("User siebel email for consider - " + email);
                        }
                    }
                    //if (log.IsDebugEnabled) log.Debug("Checking changes in Exchange");
                    if (bAllOK)
                    {
                        foreach (SEIUser u in allSettings.SEIUsers)
                        {
                            if (u.Direction == "Off") continue;
                            //if (SEI.ExchangeWSW.CheckUserChanged(u.Email, u.SyncToken) || conUsers.Contains(u.Email))
                            if (SEI.ExchangeWSW.CheckUserChanged(u.Email, u.SyncToken) || conUpperUsersMail.Contains(u.Email.ToUpper()))
                            {
                                Task t = Task.Run(() => { Syncee.SyncUser(u, allSettings.SEIUsers.ToList(), allSettings.Rooms); });
                                syncTasks.Add(t);
                            }
                        }
                        if (log.IsInfoEnabled) log.Info("All number of users to consider - " + syncTasks.Count);
                        conUsers.Clear();
                        conUpperUsersMail.Clear();
                    }
                    //if (log.IsDebugEnabled) log.Debug("Checking changes in Exchange completed");
                    Task.WaitAll(syncTasks.ToArray());
                    foreach (Task t in syncTasks.Where(f => f.IsFaulted))
                    {
                        if (t.Exception != null)
                        {
                            log.Error(String.Format("Task {0} faulted: {1}", t.Id, t.Exception.Message));
                        }

                        else
                        {
                            log.Error(String.Format("Task {0} faulted: {1}", t.Id, t.Status));
                        }

                    }
                    if (log.IsInfoEnabled) log.Info("Synchronization end");
                }
                catch (Exception ex)
                {
                    log.Error("Synchronization error: " + ex.Message);
                }
                System.Threading.Monitor.Exit(timerLock);
            }
        }

        protected override void OnStop()
        {
            using (ThreadContext.Stacks["NDC"].Push("Service stoping"))
            {
                this.timer.Stop();
                this.timer = null;

                if (log.IsInfoEnabled) if (log.IsInfoEnabled) log.Info("Service stopped");
            }
        }
    }
}
