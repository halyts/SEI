using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SEI
{
    enum ResultField
    {
        LIST_USERID,
        LIST_USERLOGIN,
        LIST_EMAIL,
        REAL_USERID,
        REAL_USERLOGIN,
        REAL_EMAIL
    }

    class SiebelDBW
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        //#pragma warning disable CS0618 // Тип или член устарел
        public static void UpdateDeletedStatus(string Id)
        {
            Oracle.DataAccess.Client.OracleConnection oConnection = new Oracle.DataAccess.Client.OracleConnection(Properties.Settings.Default.uOracleConnectionString);
            Oracle.DataAccess.Client.OracleCommand oCommand;

            try { 
                oConnection.Open();
                oCommand = new Oracle.DataAccess.Client.OracleCommand
                {
                    BindByName = true,
                    Connection = oConnection,
                    CommandType = System.Data.CommandType.Text,
                    CommandText = @"update " + Properties.Settings.Default.OracleSchema + @".cx_sei_commands c
                                         set c.command_executed = 'Y', last_upd = :2, last_upd_by = 'SEI', db_last_upd = :2, db_last_upd_src = 'SEI' where c.action_id = :1"
                };
                oCommand.Parameters.Add("1", Oracle.DataAccess.Client.OracleDbType.Varchar2);
                oCommand.Parameters["1"].Value = Id;
                oCommand.Parameters.Add("2", Oracle.DataAccess.Client.OracleDbType.Date);
                oCommand.Parameters["2"].Value = DateTime.UtcNow;
                oCommand.ExecuteNonQuery();
            }
            finally
            {
                oConnection.Close();
                oConnection.Dispose();
            }
        }
        public static void ConfirmCommand(string Id)
        {
            Oracle.DataAccess.Client.OracleConnection oConnection = new Oracle.DataAccess.Client.OracleConnection(Properties.Settings.Default.uOracleConnectionString);
            Oracle.DataAccess.Client.OracleCommand oCommand;

            try
            { 
                oConnection.Open();
                oCommand = new Oracle.DataAccess.Client.OracleCommand
                {
                    BindByName = true,
                    Connection = oConnection,
                    CommandType = System.Data.CommandType.Text,
                    CommandText = @"update " + Properties.Settings.Default.OracleSchema + @".cx_sei_commands c
                                         set c.command_executed = 'Y', last_upd = :2, last_upd_by = 'SEI', db_last_upd = :2, db_last_upd_src = 'SEI' where c.row_id = :1"
                };
                oCommand.Parameters.Add("1", Oracle.DataAccess.Client.OracleDbType.Varchar2);
                oCommand.Parameters["1"].Value = Id;
                oCommand.Parameters.Add("2", Oracle.DataAccess.Client.OracleDbType.Date);
                oCommand.Parameters["2"].Value = DateTime.UtcNow;
                oCommand.ExecuteNonQuery();
            }
            finally
            {
                oConnection.Close();
                oConnection.Dispose();
            }
        }

        public static DateTime GetCurrentDT()
        {
            Oracle.DataAccess.Client.OracleConnection oConnection = new Oracle.DataAccess.Client.OracleConnection(Properties.Settings.Default.uOracleConnectionString);
            DateTime res;

            Oracle.DataAccess.Client.OracleCommand oCommand;
            Oracle.DataAccess.Client.OracleDataReader oReader;
            try { 

                oConnection.Open();
                oCommand = new Oracle.DataAccess.Client.OracleCommand
                {
                    BindByName = true,
                    Connection = oConnection,
                    CommandType = System.Data.CommandType.Text,
                    CommandText = @"select sysdate from dual"
                };

                oReader = oCommand.ExecuteReader();
                oReader.Read();
                res = oReader.GetDateTime(0);

                return res;
            }
            finally
            {
                oConnection.Close();
                oConnection.Dispose();
            }
        }

        public static void SetResponse(string ActionExchId, DateTime lastSyncDate, string email, string response)
        {
            string sResponse = "";
            switch (response)
            {
                case "Tentative":
                    sResponse = "Tentative";
                    break;
                case "Decline":
                    sResponse = "Declined";
                    break;
                case "Accept":
                    sResponse = "Accepted";
                    break;
                default:
                    break;
            }
            if (sResponse == "")
                throw (new ApplicationException(string.Format("Unknown response type: %1, %2, %3", ActionExchId, email, response)));

            Oracle.DataAccess.Client.OracleConnection oConnection = new Oracle.DataAccess.Client.OracleConnection(Properties.Settings.Default.uOracleConnectionString);
            Oracle.DataAccess.Client.OracleCommand oCommand;
            Oracle.DataAccess.Client.OracleDataReader oReader;

            try
            {
                oConnection.Open();

                oCommand = new Oracle.DataAccess.Client.OracleCommand
                {
                    BindByName = true,
                    Connection = oConnection,
                    CommandType = System.Data.CommandType.Text,
                    CommandText = @"select row_id from " + Properties.Settings.Default.OracleSchema + @".s_user where login = :1"
                };
                oCommand.Parameters.Add("1", Oracle.DataAccess.Client.OracleDbType.Varchar2);
                oCommand.Parameters["1"].Value = Properties.Settings.Default.SiebelLogin;
                oReader = oCommand.ExecuteReader();
                if (oReader.Read())
                {
                    string tech_user_RID = oReader.GetString(0);
                }
                else
                    throw new ApplicationException("Error: " + Properties.Settings.Default.OracleLogin + " is not a siebel user");

                oCommand = new Oracle.DataAccess.Client.OracleCommand
                {
                    BindByName = true,
                    Connection = oConnection,
                    CommandType = System.Data.CommandType.Text,
                    CommandText = @"select row_id from " + Properties.Settings.Default.OracleSchema + @".s_contact c where c.email_addr = lower(:1)"
                };
                oCommand.Parameters.Add("1", Oracle.DataAccess.Client.OracleDbType.Varchar2);
                oCommand.Parameters["1"].Value = email;
                int records = oCommand.ExecuteNonQuery();
                if (records >= 0)
                {
                    oCommand = new Oracle.DataAccess.Client.OracleCommand
                    {
                        BindByName = true,
                        Connection = oConnection,
                        CommandType = System.Data.CommandType.Text,
                        CommandText = @"update " + Properties.Settings.Default.OracleSchema + @".s_act_emp ae
                                       set ae.act_invt_resp_cd = :4, ae.x_change_status_dt = :3, ae.last_upd = sys_extract_utc(systimestamp), ae.db_last_upd = sys_extract_utc(systimestamp), ae.db_last_upd_src='SEI'
                                     where ae.emp_id = (select row_id from " + Properties.Settings.Default.OracleSchema + @".s_contact c where c.email_addr = lower(:2))
                                       and ae.activity_id = (select row_id from " + Properties.Settings.Default.OracleSchema + @".cx_sei_act_x where exc_id = :1)
                                       and (ae.x_change_status_dt <= :3 or ae.x_change_status_dt is null)"
                    };
                    oCommand.Parameters.Add("1", Oracle.DataAccess.Client.OracleDbType.Varchar2);
                    oCommand.Parameters["1"].Value = ActionExchId;
                    oCommand.Parameters.Add("2", Oracle.DataAccess.Client.OracleDbType.Varchar2);
                    oCommand.Parameters["2"].Value = email;
                    oCommand.Parameters.Add("3", Oracle.DataAccess.Client.OracleDbType.Date);
                    oCommand.Parameters["3"].Value = lastSyncDate.ToUniversalTime();
                    oCommand.Parameters.Add("4", Oracle.DataAccess.Client.OracleDbType.Varchar2);
                    oCommand.Parameters["4"].Value = sResponse;

                    records = oCommand.ExecuteNonQuery();
                    if (records == 0)
                        throw (new ApplicationException("Conflict saving user response - delayed to next sync cycle"));
                }
            }
            finally
            {
                oConnection.Close();
                oConnection.Dispose();
            }
        }
        public static List<SiebelAppointment> GetDeleted()
        {
            Oracle.DataAccess.Client.OracleConnection oConnection = new Oracle.DataAccess.Client.OracleConnection(Properties.Settings.Default.uOracleConnectionStringS);
            List<SiebelAppointment> res = new List<SiebelAppointment>();

            Oracle.DataAccess.Client.OracleCommand oCommand;
            Oracle.DataAccess.Client.OracleDataReader oReader;
            try
            {

                oConnection.Close();
                oConnection.Open();
                oCommand = new Oracle.DataAccess.Client.OracleCommand
                {
                    BindByName = true,
                    Connection = oConnection,
                    CommandType = System.Data.CommandType.Text,
                    CommandText = @"select c.action_id, c.exch_id, c.exc_owner_appt_id
                                      from " + Properties.Settings.Default.OracleSchema + @".cx_sei_commands c
                                     where c.command_executed = 'N'
                                       and c.command = 'DELETE_APPT'
                                       and c.exc_owner_appt_id is not null"
                };
                oReader = oCommand.ExecuteReader();
                while (oReader.Read())
                {
                    res.Add(new SiebelAppointment(oReader.GetString(0), oReader.IsDBNull(1) ? "" : oReader.GetString(1), oReader.IsDBNull(2) ? "" : oReader.GetString(2)));
                }

                return res;
            }

            catch (Exception e)
            {
                log.Error(e.Message);
                return res;
            }
            finally
            {
                oConnection.Close();
                oConnection.Dispose();
            }
        }
        public static List<EAppointmentResponse> LoadUserResponses(SEIUser user)
        {
            List<EAppointmentResponse> res = new List<EAppointmentResponse>();
            EAppointmentResponse eAR;
            Oracle.DataAccess.Client.OracleConnection oConnection = new Oracle.DataAccess.Client.OracleConnection(Properties.Settings.Default.uOracleConnectionString);
            Oracle.DataAccess.Client.OracleCommand oCommand;
            Oracle.DataAccess.Client.OracleDataReader oReader;
            try
            { 
                oConnection.Open();

                oCommand = new Oracle.DataAccess.Client.OracleCommand
                {
                    BindByName = true,
                    Connection = oConnection,
                    CommandType = System.Data.CommandType.Text,
                    CommandText = @"select row_id from " + Properties.Settings.Default.OracleSchema + @".s_user where login = :1"
                };
                oCommand.Parameters.Add("1", Oracle.DataAccess.Client.OracleDbType.Varchar2);
                oCommand.Parameters["1"].Value = Properties.Settings.Default.SiebelLogin;
                oReader = oCommand.ExecuteReader();
                if (oReader.Read())
                {
                    string tech_user_RID = oReader.GetString(0);
                }
                else
                    throw new ApplicationException("Error: " + Properties.Settings.Default.OracleLogin + " is not a siebel user");

                oCommand = new Oracle.DataAccess.Client.OracleCommand
                {
                    BindByName = true,
                    Connection = oConnection,
                    CommandType = System.Data.CommandType.Text,
                    CommandText = @"select sa.exc_id, ae.last_upd, ae.x_change_status_dt, ae.act_invt_resp_cd
                                          from " + Properties.Settings.Default.OracleSchema + @".s_evt_act a
                                          join " + Properties.Settings.Default.OracleSchema + @".cx_sei_act_x sa
                                            on sa.par_row_id = a.row_id
                                          join " + Properties.Settings.Default.OracleSchema + @".s_act_emp ae
                                            on ae.activity_id = a.row_id
                                         where ae.x_change_status_dt > :3
                                       
                                           and ae.emp_id = :2
                                           and sa.exc_id is not null
                                           and ae.x_change_status_dt is not null"
                };
                //and ae.last_upd_by != :1
                //oCommand.Parameters.Add("1", OracleType.VarChar);
                //oCommand.Parameters["1"].Value = tech_user_RID;
                oCommand.Parameters.Add("2", Oracle.DataAccess.Client.OracleDbType.Varchar2);
                oCommand.Parameters["2"].Value = user.Id;
                oCommand.Parameters.Add("3", Oracle.DataAccess.Client.OracleDbType.Date);
                oCommand.Parameters["3"].Value = user.LastSyncDate.ToUniversalTime();

                oReader = oCommand.ExecuteReader();
                while (oReader.Read())
                {
                    eAR = new EAppointmentResponse(oReader.GetString(0), oReader.GetString(3), oReader.GetDateTime(2).ToLocalTime(), user.Email, "Siebel");
                    res.Add(eAR);
                }

                return res;
            }
            finally
            {
                oConnection.Close();
                oConnection.Dispose();
            }
        }

        public static List<SiebelAppointment> LoadUserAppointments(SEIUser user)
        {
                List<SiebelAppointment> res = new List<SiebelAppointment>();
                SiebelAppointment sA;
            Oracle.DataAccess.Client.OracleConnection oConnection = new Oracle.DataAccess.Client.OracleConnection(Properties.Settings.Default.uOracleConnectionString);
            Oracle.DataAccess.Client.OracleCommand oCommand;
                Oracle.DataAccess.Client.OracleDataReader oReader;
                try
                {
                    oConnection.Open();
                oCommand = new Oracle.DataAccess.Client.OracleCommand
                {
                    BindByName = true,
                    Connection = oConnection,
                    CommandType = System.Data.CommandType.Text,
                    CommandText = @"select row_id from " + Properties.Settings.Default.OracleSchema + @".s_user where login = :1"
                };
                oCommand.Parameters.Add("1", Oracle.DataAccess.Client.OracleDbType.Varchar2);
                    oCommand.Parameters["1"].Value = Properties.Settings.Default.SiebelLogin;
                    oReader = oCommand.ExecuteReader();
                string tech_user_RID;
                if (oReader.Read())
                {
                    tech_user_RID = oReader.GetString(0);
                }
                else
                {
                    log.Error(Properties.Settings.Default.OracleLogin + " is not a siebel user");
                    throw new ApplicationException("Error: " + Properties.Settings.Default.OracleLogin + " is not a siebel user");
                }
                oCommand = new Oracle.DataAccess.Client.OracleCommand
                {
                    BindByName = true,
                    Connection = oConnection,
                    CommandType = System.Data.CommandType.Text,
                    CommandText = @"select a.row_id,
                                                sa.exc_id,
                                                a.name,
                                                a.comments_long,
                                                a.TODO_PLAN_START_DT,
                                                a.TODO_PLAN_END_DT,
                                                x.when,
                                                sysdate sync_date,
                                                a.OWNER_PER_ID,
                                                c.last_name || ' ' || c.fst_name || case
                                                    when c.mid_name is not null then
                                                    ' ' || c.mid_name
                                                    else
                                                    null
                                                end,
                                                c.email_addr owner_email,
                                                sa.exc_owner_appt_id,
                                                a.LOC_DESC,
                                                a.APPT_ALARM_TM_MIN
                                            from " + Properties.Settings.Default.OracleSchema + @".s_evt_act a
                                            left join " + Properties.Settings.Default.OracleSchema + @".cx_sei_act_x sa
                                            on sa.par_row_id = a.row_id
                                            join " + Properties.Settings.Default.OracleSchema + @".s_contact c
                                            on c.row_id = a.owner_per_id
                                            join (select distinct row_id, when
                                                    from (select /*distinct*/
                                                            row_id,
                                                            max(last_upd_by) keep(dense_rank last order by last_upd) over(partition by row_id) who,
                                                            max(last_upd) over(partition by row_id) when
                                                            from (select a.row_id, a.last_upd, a.last_upd_by
                                                                    from " + Properties.Settings.Default.OracleSchema + @".s_evt_act a
                                                                    join " + Properties.Settings.Default.OracleSchema + @".s_lst_of_val lv on lv.type ='ACTIVITY_DISPLAY_CODE' and lv.name = 'Calendar and Activity' and lv.lang_id = 'RUS' and lv.val = a.cal_type_cd
                                                                    left join " + Properties.Settings.Default.OracleSchema + @".cx_sei_act_x sa
                                                                    on sa.par_row_id = a.row_id
                                                                    where a.OWNER_PER_ID = :2
                                                                    /*and a.cal_type_cd =
                                                                        (select val from s_lst_of_val where type ='ACTIVITY_DISPLAY_CODE' and name = 'Calendar and Activity' and lang_id = 'RUS')*/
                                                                    and (sa.sync_status is not null and
                                                                        sa.sync_status not in ('ED', 'CR','ExEtS'))
                                                                union all
                                                                select a.row_id, ae.last_upd, ae.last_upd_by
                                                                    from " + Properties.Settings.Default.OracleSchema + @".s_evt_act a
                                                                    join " + Properties.Settings.Default.OracleSchema + @".s_lst_of_val lv
                                                                    on lv.type ='ACTIVITY_DISPLAY_CODE' and lv.name = 'Calendar and Activity' and lv.lang_id = 'RUS' and lv.val = a.cal_type_cd
                                                                    left join " + Properties.Settings.Default.OracleSchema + @".cx_sei_act_x sa
                                                                    on sa.par_row_id = a.row_id
                                                                    join " + Properties.Settings.Default.OracleSchema + @".s_act_emp ae
                                                                    on ae.activity_id = a.row_id
                                                                    where a.OWNER_PER_ID = :2
                                                                    /*and a.cal_type_cd =
                                                                        (select val from s_lst_of_val where type ='ACTIVITY_DISPLAY_CODE' and name = 'Calendar and Activity' and lang_id = 'RUS')*/
                                                                    and (sa.sync_status is not null and
                                                                        sa.sync_status not in ('ED', 'CR','ExEtS'))
                                                                union all
                                                                select a.row_id, ac.last_upd, ac.last_upd_by
                                                                    from " + Properties.Settings.Default.OracleSchema + @".s_evt_act a
                                                                    join " + Properties.Settings.Default.OracleSchema + @".s_lst_of_val lv
                                                                    on lv.type ='ACTIVITY_DISPLAY_CODE' and lv.name = 'Calendar and Activity' and lv.lang_id = 'RUS' and lv.val = a.cal_type_cd
                                                                    left join " + Properties.Settings.Default.OracleSchema + @".cx_sei_act_x sa
                                                                    on sa.par_row_id = a.row_id
                                                                    join " + Properties.Settings.Default.OracleSchema + @".s_act_contact ac
                                                                    on ac.activity_id = a.row_id
                                                                    where a.OWNER_PER_ID = :2
                                                                    /*and a.cal_type_cd =
                                                                        (select val from s_lst_of_val where type ='ACTIVITY_DISPLAY_CODE' and name = 'Calendar and Activity' and lang_id = 'RUS')*/
                                                                    and (sa.sync_status is not null and
                                                                        sa.sync_status not in ('ED', 'CR','ExEtS'))
                                                                union all
                                                                select a.row_id, ar.last_upd, ar.last_upd_by
                                                                    from " + Properties.Settings.Default.OracleSchema + @".s_evt_act a
                                                                    join " + Properties.Settings.Default.OracleSchema + @".s_lst_of_val lv
                                                                    on lv.type ='ACTIVITY_DISPLAY_CODE' and lv.name = 'Calendar and Activity' and lv.lang_id = 'RUS' and lv.val = a.cal_type_cd
                                                                    left join " + Properties.Settings.Default.OracleSchema + @".cx_sei_act_x sa
                                                                    on sa.par_row_id = a.row_id
                                                                    join " + Properties.Settings.Default.OracleSchema + @".S_ACT_CAL_RSRC ar
                                                                    on ar.activity_id = a.row_id
                                                                    where a.OWNER_PER_ID = :2
                                                                    /*and a.cal_type_cd =
                                                                        (select val from s_lst_of_val where type ='ACTIVITY_DISPLAY_CODE' and name = 'Calendar and Activity' and lang_id = 'RUS')*/
                                                                    and (sa.sync_status is not null and
                                                                        sa.sync_status not in ('ED', 'CR','ExEtS'))
                                                                union all
                                                                select a.row_id, sc.last_upd, sc.last_upd_by
                                                                    from " + Properties.Settings.Default.OracleSchema + @".s_evt_act a
                                                                    join " + Properties.Settings.Default.OracleSchema + @".s_lst_of_val lv
                                                                    on lv.type ='ACTIVITY_DISPLAY_CODE' and lv.name = 'Calendar and Activity' and lv.lang_id = 'RUS' and lv.val = a.cal_type_cd
                                                                    left join " + Properties.Settings.Default.OracleSchema + @".cx_sei_act_x sa
                                                                    on sa.par_row_id = a.row_id
                                                                    join " + Properties.Settings.Default.OracleSchema + @".cx_sei_commands sc
                                                                    on sc.command in
                                                                        ('DELETE_EMP', 'DELETE_CON', 'DELETE_RES')
                                                                    and sc.action_id = a.row_id
                                                                    where a.OWNER_PER_ID = :2
                                                                    /*and a.cal_type_cd =
                                                                        (select val from s_lst_of_val where type ='ACTIVITY_DISPLAY_CODE' and name = 'Calendar and Activity' and lang_id = 'RUS')*/
                                                                    and (sa.sync_status is not null and
                                                                        sa.sync_status not in ('ED', 'CR', 'ExEtS'))))
                                                    where who != :1
                                                    and when >= :3) x
                                            on x.row_id = a.row_id"
                };
                oCommand.Parameters.Add("1", Oracle.DataAccess.Client.OracleDbType.Varchar2);
                    oCommand.Parameters["1"].Value = tech_user_RID;
                    oCommand.Parameters.Add("2", Oracle.DataAccess.Client.OracleDbType.Varchar2);
                    oCommand.Parameters["2"].Value = user.Id;
                    oCommand.Parameters.Add("3", Oracle.DataAccess.Client.OracleDbType.Date);
                    oCommand.Parameters["3"].Value = user.LastSyncDate.ToUniversalTime();

                    oReader = oCommand.ExecuteReader();
                    while (oReader.Read())
                    {
                        sA = new SiebelAppointment(oReader.GetString(0), oReader.IsDBNull(1) ? "" : oReader.GetString(1), oReader.IsDBNull(2) ? "" : oReader.GetString(2), oReader.IsDBNull(3) ? "" : oReader.GetString(3), oReader.GetDateTime(4).ToLocalTime(), oReader.GetDateTime(5).ToLocalTime(), oReader.GetDateTime(6).ToLocalTime(), oReader.GetDateTime(7), oReader.IsDBNull(11) ? "" : oReader.GetString(11)
                            , oReader.IsDBNull(12) ? "" : oReader.GetString(12)
                            , oReader.IsDBNull(13) ? 0 : Convert.ToInt32(oReader.GetDecimal(13)));
                        sA.Owner.Id = oReader.GetString(8);
                        sA.Owner.Name = oReader.GetString(9);
                        sA.Owner.Email = oReader.IsDBNull(10) ? "" : oReader.GetString(10);
                        res.Add(sA);
                    }
                oCommand = new Oracle.DataAccess.Client.OracleCommand
                {
                    BindByName = true,
                    Connection = oConnection,
                    CommandType = System.Data.CommandType.Text,
                    CommandText = @"select distinct e.emp_id,
                                                c.last_name || ' ' || c.fst_name || case
                                                    when c.mid_name is not null then
                                                    ' ' || c.mid_name
                                                    else
                                                    null
                                                end,
                                                c.email_addr,
                                                e.act_invt_resp_cd
                                            from " + Properties.Settings.Default.OracleSchema + @".s_act_emp e
                                            join " + Properties.Settings.Default.OracleSchema + @".s_contact c
                                            on c.row_id = e.emp_id
                                            where e.activity_id = :1
                                            and e.emp_id != :2
                                            and e.x_sei_is_required = 'Y'
                                            and c.email_addr is not null"
                };
                oCommand.Parameters.Add("1", Oracle.DataAccess.Client.OracleDbType.Varchar2);
                oCommand.Parameters.Add("2", Oracle.DataAccess.Client.OracleDbType.Varchar2);

                    for (int i = 0; i < res.Count; i++)
                    {
                        oCommand.Parameters["1"].Value = res[i].SiebelId;
                        oCommand.Parameters["2"].Value = res[i].Owner.Id;
                        oReader = oCommand.ExecuteReader();
                        while (oReader.Read())
                        {
                            res[i].RequiredAttendees.Add(new SiebelAttendee(oReader.GetString(0), oReader.GetString(1), oReader.IsDBNull(2) ? "" : oReader.GetString(2), oReader.IsDBNull(3) ? "Unknown" : oReader.GetString(3)));
                        }
                    }

                    oCommand.CommandText = @"select distinct x.con_id,
                                                              c.last_name || ' ' || c.fst_name || case
                                                                  when c.mid_name is not null then
                                                                  ' ' || c.mid_name
                                                                  else
                                                                  null
                                                              end,
                                                              cd.value
                                                          from " + Properties.Settings.Default.OracleSchema + @".s_act_contact x
                                                          join " + Properties.Settings.Default.OracleSchema + @".s_contact c on c.row_id = x.con_id
                                                          join " + Properties.Settings.Default.OracleSchema + @".cx_contact_data cd on cd.par_row_id = c.row_id and cd.active = 'Y' and cd.contact_type = 'Основной Email'
                                                          where x.activity_id = :1
                                                          and x.con_id != :2
                                                          and x.x_sei_is_required = 'Y' 
                                                          and cd.value is not null";

                    //VG

                    //oCommand.CommandText = @"select distinct x.con_id,
                    //                                         c.last_name || ' ' || c.fst_name || case
                    //                                             when c.mid_name is not null then
                    //                                             ' ' || c.mid_name
                    //                                             else
                    //                                             null
                    //                                         end,
                    //                                         nvl(cd.value, 'ihave@not.mail') value 
                    //                                     from " + Properties.Settings.Default.OracleSchema + @".s_act_contact x
                    //                                     join " + Properties.Settings.Default.OracleSchema + @".s_contact c on c.row_id = x.con_id
                    //                                     left join " + Properties.Settings.Default.OracleSchema + @".cx_contact_data cd on cd.par_row_id = c.row_id and cd.active = 'Y' and cd.contact_type = 'Основной Email'
                    //                                     where x.activity_id = :1
                    //                                     and x.con_id != :2
                    //                                     and x.x_sei_is_required = 'Y'";

                    //VG
                    for (int i = 0; i < res.Count; i++)
                    {
                        oCommand.Parameters["1"].Value = res[i].SiebelId;
                        oCommand.Parameters["2"].Value = res[i].Owner.Id;
                        oReader = oCommand.ExecuteReader();
                        while (oReader.Read())
                        {
                            res[i].RequiredAttendees.Add(new SiebelAttendee(oReader.GetString(0), oReader.GetString(1), oReader.IsDBNull(2) ? "" : oReader.GetString(2), "Unknown"));
                        }
                    }

                    oCommand.CommandText = @"select distinct e.emp_id,
                                                c.last_name || ' ' || c.fst_name || case
                                                    when c.mid_name is not null then
                                                    ' ' || c.mid_name
                                                    else
                                                    null
                                                end,
                                                c.email_addr,
                                                e.act_invt_resp_cd
                                            from " + Properties.Settings.Default.OracleSchema + @".s_act_emp e
                                            join " + Properties.Settings.Default.OracleSchema + @".s_contact c
                                            on c.row_id = e.emp_id
                                            where e.activity_id = :1
                                            and e.emp_id != :2
                                            and(e.x_sei_is_required = 'N' OR e.x_sei_is_required is null)
                                            and c.email_addr is not null";

                    for (int i = 0; i < res.Count; i++)
                    {
                        oCommand.Parameters["1"].Value = res[i].SiebelId;
                        oCommand.Parameters["2"].Value = res[i].Owner.Id;
                        oReader = oCommand.ExecuteReader();
                        while (oReader.Read())
                        {
                            res[i].OptionalAttendees.Add(new SiebelAttendee(oReader.GetString(0), oReader.GetString(1), oReader.IsDBNull(2) ? "" : oReader.GetString(2), oReader.IsDBNull(3) ? "Unknown" : oReader.GetString(3)));
                        }
                    }

                    oCommand.CommandText = @"select distinct x.con_id,
                                          c.last_name || ' ' || c.fst_name || case
                                              when c.mid_name is not null then
                                              ' ' || c.mid_name
                                              else
                                              null
                                          end,
                                          cd.value from " + Properties.Settings.Default.OracleSchema + @".s_act_contact x
                                      join " + Properties.Settings.Default.OracleSchema + @".s_contact c on c.row_id = x.con_id
                                      join " + Properties.Settings.Default.OracleSchema + @".cx_contact_data cd on cd.par_row_id = c.row_id and cd.active = 'Y' and cd.contact_type = 'Основной Email'
                                      where x.activity_id = :1
                                      and x.con_id != :2
                                      and (x.x_sei_is_required = 'N' OR x.x_sei_is_required is null)
                                      and cd.value is not null";
                    for (int i = 0; i < res.Count; i++)
                    {
                        oCommand.Parameters["1"].Value = res[i].SiebelId;
                        oCommand.Parameters["2"].Value = res[i].Owner.Id;
                        oReader = oCommand.ExecuteReader();
                        while (oReader.Read())
                        {
                            res[i].OptionalAttendees.Add(new SiebelAttendee(oReader.GetString(0), oReader.GetString(1), oReader.IsDBNull(2) ? "" : oReader.GetString(2), "Unknown"));
                        }
                    }

                    oCommand.CommandText = @"select distinct r.row_id, r.name, r.x_sei_res_addr
                                          from " + Properties.Settings.Default.OracleSchema + @".S_ACT_CAL_RSRC ar
                                          join " + Properties.Settings.Default.OracleSchema + @".s_cal_rsrc r
                                            on r.row_id = ar.resource_id
                                         where ar.activity_id = :1";
                    oCommand.Parameters.Remove(oCommand.Parameters["2"]);
                    for (int i = 0; i < res.Count; i++)
                    {
                        oCommand.Parameters["1"].Value = res[i].SiebelId;
                        oReader = oCommand.ExecuteReader();
                        while (oReader.Read())
                        {
                            res[i].Resources.Add(new SiebelAttendee(oReader.GetString(0), oReader.IsDBNull(1) ? "" : oReader.GetString(1), oReader.IsDBNull(2) ? "" : oReader.GetString(2), "Unknown"));
                        }
                    }

                    return res;
                }
                finally
                {                
                    oConnection.Close();
                    oConnection.Dispose();
                }
            
        
        }

        public static void SaveUserSyncState(string userLogin, DateTime syncDate, string syncToken)
        {
            Oracle.DataAccess.Client.OracleConnection oConnection = new Oracle.DataAccess.Client.OracleConnection(Properties.Settings.Default.uOracleConnectionString);
            Oracle.DataAccess.Client.OracleCommand oCommand;

            try
            {
                oConnection.Open();
                oCommand = new Oracle.DataAccess.Client.OracleCommand
                {
                    BindByName = true,
                    Connection = oConnection,
                    CommandType = System.Data.CommandType.Text,
                    //oCommand.CommandText = @"update " + Properties.Settings.Default.OracleSchema + @".cx_sei_users u set u.last_sync_date = :1, u.sync_token_l = :2, u.last_error = '' where u.user_login = :3";
                    //My
                    CommandText = @"update " + Properties.Settings.Default.OracleSchema + @".cx_sei_users u set u.last_sync_date = :1, u.sync_token = :2, u.last_error = '' where u.user_login = :3"
                };

                oCommand.Parameters.Add("1", Oracle.DataAccess.Client.OracleDbType.Date);
                oCommand.Parameters["1"].Value = syncDate.ToUniversalTime();
                oCommand.Parameters.Add("2", Oracle.DataAccess.Client.OracleDbType.Varchar2);
                oCommand.Parameters["2"].Value = syncToken;
                oCommand.Parameters.Add("3", Oracle.DataAccess.Client.OracleDbType.Varchar2);
                oCommand.Parameters["3"].Value = userLogin;
                int records = oCommand.ExecuteNonQuery();
                if (records == 0)
                    throw (new ApplicationException("Cannot save usersyncstate"));
            }
            finally
            {
                oConnection.Close();
                oConnection.Dispose();
            }
        }
        public static void SaveUserSyncState(string userLogin, string syncToken)
        {
            Oracle.DataAccess.Client.OracleConnection oConnection = new Oracle.DataAccess.Client.OracleConnection(Properties.Settings.Default.uOracleConnectionString);
            Oracle.DataAccess.Client.OracleCommand oCommand;

            try
            {
                oConnection.Open();
                oCommand = new Oracle.DataAccess.Client.OracleCommand
                {
                    BindByName = true,
                    Connection = oConnection,
                    CommandType = System.Data.CommandType.Text,
                    //oCommand.CommandText = @"update " + Properties.Settings.Default.OracleSchema + @".cx_sei_users u set u.sync_token_l = :2, u.last_error = '' where u.user_login = :3";
                    //My
                    CommandText = @"update " + Properties.Settings.Default.OracleSchema + @".cx_sei_users u set u.sync_token = :2, u.last_error = '' where u.user_login = :3"
                };
                oCommand.Parameters.Add("2", Oracle.DataAccess.Client.OracleDbType.Varchar2);
                oCommand.Parameters["2"].Value = syncToken;
                oCommand.Parameters.Add("3", Oracle.DataAccess.Client.OracleDbType.Varchar2);
                oCommand.Parameters["3"].Value = userLogin;
                int records = oCommand.ExecuteNonQuery();
                if (records == 0)
                    throw (new ApplicationException("Cannot save usersyncstate"));
            }
            finally
            {
                oConnection.Close();
                oConnection.Dispose();
            }
        }

        public static void SaveUserError(string userLogin, string errorText)
        {
            Oracle.DataAccess.Client.OracleConnection oConnection = new Oracle.DataAccess.Client.OracleConnection(Properties.Settings.Default.uOracleConnectionString);
            Oracle.DataAccess.Client.OracleCommand oCommand;

            try
            {
                oConnection.Open();
                oCommand = new Oracle.DataAccess.Client.OracleCommand
                {
                    BindByName = true,
                    Connection = oConnection,
                    CommandType = System.Data.CommandType.Text,
                    CommandText = @"update " + Properties.Settings.Default.OracleSchema + @".cx_sei_users u set u.last_error = :2 where u.user_login = :1"
                };
                oCommand.Parameters.Add("1", Oracle.DataAccess.Client.OracleDbType.Varchar2);
                oCommand.Parameters["1"].Value = userLogin;
                oCommand.Parameters.Add("2", Oracle.DataAccess.Client.OracleDbType.Varchar2);
                oCommand.Parameters["2"].Value = errorText;
                oCommand.ExecuteNonQuery();
            }
            finally
            {
                oConnection.Close();
                oConnection.Dispose();
            }
}
        public static void SaveAppointmentId(SiebelAppointment appointment)
        {
            Oracle.DataAccess.Client.OracleConnection oConnection = new Oracle.DataAccess.Client.OracleConnection(Properties.Settings.Default.uOracleConnectionString);
            Oracle.DataAccess.Client.OracleCommand oCommand;

            try
            {
                oConnection.Open();

                oCommand = new Oracle.DataAccess.Client.OracleCommand
                {
                    BindByName = true,
                    Connection = oConnection,
                    CommandType = System.Data.CommandType.Text,
                    CommandText = @"update " + Properties.Settings.Default.OracleSchema + @".cx_sei_act_x a set a.exc_id = :2, a.exc_owner_appt_id = :3 where a.row_id = :1"
                };
                oCommand.Parameters.Add("1", Oracle.DataAccess.Client.OracleDbType.Varchar2);
                oCommand.Parameters["1"].Value = appointment.SiebelId;
                oCommand.Parameters.Add("2", Oracle.DataAccess.Client.OracleDbType.Varchar2);
                oCommand.Parameters["2"].Value = appointment.Id;
                oCommand.Parameters.Add("3", Oracle.DataAccess.Client.OracleDbType.Varchar2);
                oCommand.Parameters["3"].Value = appointment.OwnerId;
                int records = oCommand.ExecuteNonQuery();
                if (records == 0)
                {
                    oCommand.CommandText = @"insert into " + Properties.Settings.Default.OracleSchema + @".cx_sei_act_x
                                                (row_id,
                                                created,
                                                created_by,
                                                last_upd,
                                                last_upd_by,
                                                modification_num,
                                                conflict_id,
                                                par_row_id,
                                                db_last_upd,
                                                db_last_upd_src,
                                                exc_id,
                                                exc_owner_appt_id)
                                            values
                                                (:1,
                                                sys_extract_utc(systimestamp),
                                                (select row_id from " + Properties.Settings.Default.OracleSchema + @".s_user where login = :4),
                                                sys_extract_utc(systimestamp),
                                                (select row_id from " + Properties.Settings.Default.OracleSchema + @".s_user where login = :4),
                                                0,
                                                0,
                                                :1,
                                                sys_extract_utc(systimestamp),
                                                'SEI',
                                                :2,
                                                :3)";
                    oCommand.Parameters.Add("4", Oracle.DataAccess.Client.OracleDbType.Varchar2);
                    oCommand.Parameters["4"].Value = Properties.Settings.Default.OracleLogin;
                    oCommand.ExecuteNonQuery();
                }
            }
            finally
            {
                oConnection.Close();
                oConnection.Dispose();
            }
        }
        public static void SaveAppointmentSyncState(SiebelAppointment appointment, string sState)
        {
            Oracle.DataAccess.Client.OracleConnection oConnection = new Oracle.DataAccess.Client.OracleConnection(Properties.Settings.Default.uOracleConnectionString);
            Oracle.DataAccess.Client.OracleCommand oCommand;

            try
            {
                oConnection.Open();
                oCommand = new Oracle.DataAccess.Client.OracleCommand
                {
                    BindByName = true,
                    Connection = oConnection,
                    CommandType = System.Data.CommandType.Text,
                    CommandText = @"update " + Properties.Settings.Default.OracleSchema + @".cx_sei_act_x a set a.sync_status = :2 where a.row_id = :1"
                };
                oCommand.Parameters.Add("1", Oracle.DataAccess.Client.OracleDbType.Varchar2);
                oCommand.Parameters["1"].Value = appointment.SiebelId;
                oCommand.Parameters.Add("2", Oracle.DataAccess.Client.OracleDbType.Varchar2);
                oCommand.Parameters["2"].Value = sState;
                int records = oCommand.ExecuteNonQuery();
                if (records == 0)
                    throw new ApplicationException("Cannot save state, Action does not exist");
                oCommand.CommandText = @"update " + Properties.Settings.Default.OracleSchema + @".s_evt_act a set a.modification_num = a.modification_num + 1 where a.row_id = :1";
                oCommand.Parameters.Remove(oCommand.Parameters["2"]);
                records = oCommand.ExecuteNonQuery();
                if (records == 0)
                    throw new ApplicationException("Cannot save state, Action does not exist");
            }
            finally
            {
                oConnection.Close();
                oConnection.Dispose();
            }
        }

        public static DateTime GetMaxDTLastSync()
        {
            DateTime res;
            Oracle.DataAccess.Client.OracleConnection oConnection = new Oracle.DataAccess.Client.OracleConnection(Properties.Settings.Default.uOracleConnectionString);
            Oracle.DataAccess.Client.OracleCommand oCommand;
            Oracle.DataAccess.Client.OracleDataReader oReader;
            try
            {
                oConnection.Open();
                oCommand = new Oracle.DataAccess.Client.OracleCommand
                {
                    BindByName = true,
                    Connection = oConnection,
                    CommandType = System.Data.CommandType.Text,
                    CommandText = @"select max(last_sync_date) from " + Properties.Settings.Default.OracleSchema + @".cx_sei_users"
                };
                oReader = oCommand.ExecuteReader();
                oReader.Read();
                res = oReader.GetDateTime(0);

            }

            catch
            {
                return DateTime.MaxValue;
            }

            finally
            {
                oConnection.Close();
                oConnection.Dispose();
            }

            return res;
        }
        public static List<string> GetConsiderUserList(DateTime dtLastSync)
        {
            if (log.IsDebugEnabled) log.Debug("Date and time of last synchronization - " + dtLastSync);
            List<string> changedUserMails = new List<string>();
            string tech_user_RID = "";

            Oracle.DataAccess.Client.OracleConnection oConnection = new Oracle.DataAccess.Client.OracleConnection(Properties.Settings.Default.uOracleConnectionString);
            Oracle.DataAccess.Client.OracleCommand oCommand;
            Oracle.DataAccess.Client.OracleDataReader oReader;
            try
            {
                oConnection.Open();
                oCommand = new Oracle.DataAccess.Client.OracleCommand
                {
                    BindByName = true,
                    Connection = oConnection,
                    CommandType = System.Data.CommandType.Text,
                    CommandText = @"select row_id from " + Properties.Settings.Default.OracleSchema + @".s_user where login = :1"
                };
                oCommand.Parameters.Add("1", Oracle.DataAccess.Client.OracleDbType.Varchar2);
                oCommand.Parameters["1"].Value = Properties.Settings.Default.OracleLogin;
                oReader = oCommand.ExecuteReader();
                if (oReader.Read())
                {
                    tech_user_RID = oReader.GetString(0);
                }
                else
                {
                    log.Error(Properties.Settings.Default.OracleLogin + " is not a siebel user");
                    throw new ApplicationException("Error: " + Properties.Settings.Default.OracleLogin + " is not a siebel user");
                }
            }

            finally
            {
                oConnection.Close();
                //oConnection.Dispose();
            }
            try
            {
                oConnection.Open();
                oCommand = new Oracle.DataAccess.Client.OracleCommand
                {
                    BindByName = true,
                    Connection = oConnection,
                    CommandType = System.Data.CommandType.Text,
                    CommandText = @"select a.row_id,
                                                sa.exc_id,
                                                a.name,
                                                a.comments_long,
                                                a.TODO_PLAN_START_DT,
                                                a.TODO_PLAN_END_DT,
                                                x.when,
                                                sysdate sync_date,
                                                a.OWNER_PER_ID,
                                                c.last_name || ' ' || c.fst_name || case
                                                    when c.mid_name is not null then
                                                    ' ' || c.mid_name
                                                    else
                                                    null
                                                end,
                                                c.email_addr owner_email,
                                                sa.exc_owner_appt_id,
                                                a.LOC_DESC,
                                                a.APPT_ALARM_TM_MIN
                                            from " + Properties.Settings.Default.OracleSchema + @".s_evt_act a
                                            left join " + Properties.Settings.Default.OracleSchema + @".cx_sei_act_x sa
                                            on sa.par_row_id = a.row_id
                                            join " + Properties.Settings.Default.OracleSchema + @".s_contact c
                                            on c.row_id = a.owner_per_id
                                            join (select distinct row_id, when
                                                    from (select 
                                                            row_id,
                                                            max(last_upd_by) keep(dense_rank last order by last_upd) over(partition by row_id) who,
                                                            max(last_upd) over(partition by row_id) when
                                                            from (select a.row_id, a.last_upd, a.last_upd_by
                                                                    from " + Properties.Settings.Default.OracleSchema + @".s_evt_act a
                                                                    join " + Properties.Settings.Default.OracleSchema + @".s_lst_of_val lv on lv.type ='ACTIVITY_DISPLAY_CODE' and lv.name = 'Calendar and Activity' and lv.lang_id = 'RUS' and lv.val = a.cal_type_cd
                                                                    left join " + Properties.Settings.Default.OracleSchema + @".cx_sei_act_x sa
                                                                    on sa.par_row_id = a.row_id
                                                                    where 
                                                                     sa.sync_status = 'RX'                                                                         
                                                                union all
                                                                select a.row_id, ae.last_upd, ae.last_upd_by
                                                                    from " + Properties.Settings.Default.OracleSchema + @".s_evt_act a
                                                                    join " + Properties.Settings.Default.OracleSchema + @".s_lst_of_val lv
                                                                    on lv.type ='ACTIVITY_DISPLAY_CODE' and lv.name = 'Calendar and Activity' and lv.lang_id = 'RUS' and lv.val = a.cal_type_cd
                                                                    left join " + Properties.Settings.Default.OracleSchema + @".cx_sei_act_x sa
                                                                    on sa.par_row_id = a.row_id
                                                                    join " + Properties.Settings.Default.OracleSchema + @".s_act_emp ae
                                                                    on ae.activity_id = a.row_id
                                                                    where 
                                                                    sa.sync_status = 'RX'
                                                                union all
                                                                select a.row_id, ac.last_upd, ac.last_upd_by
                                                                    from " + Properties.Settings.Default.OracleSchema + @".s_evt_act a
                                                                    join " + Properties.Settings.Default.OracleSchema + @".s_lst_of_val lv
                                                                    on lv.type ='ACTIVITY_DISPLAY_CODE' and lv.name = 'Calendar and Activity' and lv.lang_id = 'RUS' and lv.val = a.cal_type_cd
                                                                    left join " + Properties.Settings.Default.OracleSchema + @".cx_sei_act_x sa
                                                                    on sa.par_row_id = a.row_id
                                                                    join " + Properties.Settings.Default.OracleSchema + @".s_act_contact ac
                                                                    on ac.activity_id = a.row_id
                                                                    where 
                                                                    sa.sync_status = 'RX'
                                                                union all
                                                                select a.row_id, ar.last_upd, ar.last_upd_by
                                                                    from " + Properties.Settings.Default.OracleSchema + @".s_evt_act a
                                                                    join " + Properties.Settings.Default.OracleSchema + @".s_lst_of_val lv
                                                                    on lv.type ='ACTIVITY_DISPLAY_CODE' and lv.name = 'Calendar and Activity' and lv.lang_id = 'RUS' and lv.val = a.cal_type_cd
                                                                    left join " + Properties.Settings.Default.OracleSchema + @".cx_sei_act_x sa
                                                                    on sa.par_row_id = a.row_id
                                                                    join " + Properties.Settings.Default.OracleSchema + @".S_ACT_CAL_RSRC ar
                                                                    on ar.activity_id = a.row_id
                                                                    where 
                                                                    sa.sync_status is not null and
                                                                        sa.sync_status = 'RX' 
                                                                union all
                                                                select a.row_id, sc.last_upd, sc.last_upd_by
                                                                    from " + Properties.Settings.Default.OracleSchema + @".s_evt_act a
                                                                    join " + Properties.Settings.Default.OracleSchema + @".s_lst_of_val lv
                                                                    on lv.type ='ACTIVITY_DISPLAY_CODE' and lv.name = 'Calendar and Activity' and lv.lang_id = 'RUS' and lv.val = a.cal_type_cd
                                                                    left join " + Properties.Settings.Default.OracleSchema + @".cx_sei_act_x sa
                                                                    on sa.par_row_id = a.row_id
                                                                    join " + Properties.Settings.Default.OracleSchema + @".cx_sei_commands sc
                                                                    on sc.command in
                                                                        ('DELETE_EMP', 'DELETE_CON', 'DELETE_RES')
                                                                    and sc.action_id = a.row_id
                                                                    where 
                                                                    sa.sync_status = 'RX'
                                                                        ))
                                                    where who != :1
                                                    and when >= :3) x
                                            on x.row_id = a.row_id"
                };
                oCommand.Parameters.Add("1", Oracle.DataAccess.Client.OracleDbType.Varchar2);
                oCommand.Parameters["1"].Value = tech_user_RID;
                oCommand.Parameters.Add("3", Oracle.DataAccess.Client.OracleDbType.Date);
                oCommand.Parameters["3"].Value = dtLastSync.ToUniversalTime();
                if (log.IsDebugEnabled) log.Debug("Universal time of last synchronization - " + oCommand.Parameters["3"].Value);

                oReader = oCommand.ExecuteReader();
                while (oReader.Read())
                {
                        changedUserMails.Add(oReader.IsDBNull(10) ? "" : oReader.GetString(10));
                }

            }

            catch (Exception e)
            {
                log.Error(e.Message);
                return changedUserMails;
            }

            finally
            {
                oConnection.Close();
                oConnection.Dispose();
            }

            if (log.IsInfoEnabled) log.Info("Number of Siebel user's with changes - " + changedUserMails.Count);
            return changedUserMails;
        }
        public static void UpdateListUsers()
        {
            Oracle.DataAccess.Client.OracleConnection oConnection = new Oracle.DataAccess.Client.OracleConnection(Properties.Settings.Default.uOracleConnectionString);
            Oracle.DataAccess.Client.OracleCommand oCommand;
            Oracle.DataAccess.Client.OracleDataReader oReader;
            try
            {
                oConnection.Open();
                oCommand = new Oracle.DataAccess.Client.OracleCommand
                {
                    BindByName = true,
                    Connection = oConnection,
                    CommandType = System.Data.CommandType.Text,
                    CommandText = @"SELECT
                        user_id        list_userid,
                        user_login     list_userlogin,
                        email_addr     list_email,
                        sc_person_uid  real_userid,
                        sus_login      real_userlogin,
                        sc_email_addr  real_email
                    FROM
                        (
                            SELECT
                                user_id,
                                user_login,
                                email_addr,
                                sc_person_uid,
                                sus_login,
                                sc_email_addr
                            FROM "
                                + Properties.Settings.Default.OracleSchema + @".cx_sei_users q1
                                FULL JOIN (
                                    SELECT DISTINCT
                                        sc_person_uid,
                                        sus_login,
                                        sc_email_addr
                                    FROM
                                        (
                                            SELECT
                                                sp.party_uid     sp_party_uid,
                                                sc.person_uid    sc_person_uid,
                                                sus.login        sus_login,
                                                sc.email_addr    sc_email_addr
                                            FROM "
                                                     + Properties.Settings.Default.OracleSchema + @".s_party sp
                                                JOIN " + Properties.Settings.Default.OracleSchema + @".s_contact  sc ON sp.party_uid = sc.par_row_id
                                                JOIN " + Properties.Settings.Default.OracleSchema + @".s_user     sus ON sc.person_uid = sus.row_id
                                            WHERE
                                                    sc.emp_flg = 'Y'
                                                AND sus.x_active = 'Y'
                                                AND sc.person_uid NOT IN (
                                                    '0-1'
                                                )
                                                AND sc.person_uid NOT IN (
                                                    SELECT DISTINCT
                                                        sc_person_uid
                                                    FROM
                                                        (
                                                            SELECT
                                                                sc.person_uid sc_person_uid
                                                            FROM "
                                                                     + Properties.Settings.Default.OracleSchema + @".s_party sp
                                                                JOIN " + Properties.Settings.Default.OracleSchema + @".s_contact    sc ON sp.party_uid = sc.par_row_id
                                                                JOIN " + Properties.Settings.Default.OracleSchema + @".s_party_per  spp ON sc.person_uid = spp.person_id
                                                                JOIN " + Properties.Settings.Default.OracleSchema + @".s_postn      spn ON spp.party_id = spn.row_id
                                                            WHERE
                                                                ( sc.emp_flg = 'Y' )
                                                                AND spn.name = '00-Деактивированные пользователи'
                                                        )
                                                )
                                        )
                                    WHERE
                                        sc_email_addr IS NOT NULL
                                ) q2 ON q1.user_id = q2.sc_person_uid
                        )
                    WHERE
                        user_id IS NULL
                        OR sc_person_uid IS NULL
                        OR ( email_addr IS NOT NULL
                             AND sc_email_addr IS NOT NULL
                             AND email_addr <> sc_email_addr )"
                };

                oReader = oCommand.ExecuteReader();
                string a = string.Empty;
                string b = string.Empty;
                int cnt = 0;
                while (oReader.Read())
                {
                    cnt++;
                    if (!oReader.IsDBNull((int)ResultField.LIST_USERID) && oReader.IsDBNull((int)ResultField.REAL_USERID))
                    {
                        SiebelWSW.DeleteSEIUser(oReader.GetString((int)ResultField.LIST_USERID));
                    }
                    else
                    {
                        if (oReader.IsDBNull((int)ResultField.LIST_USERID) && !oReader.IsDBNull((int)ResultField.REAL_USERID))
                        {
                            SiebelWSW.InsertSEIUser(oReader.GetString((int)ResultField.REAL_USERID), oReader.GetString((int)ResultField.REAL_EMAIL));
                        }
                        else
                        {
                            if (log.IsWarnEnabled) log.Warn("User " + oReader.GetString((int)ResultField.REAL_USERID) + " has different emails in Siebel and SEI users list");
                        }
                    }
                }
                if (log.IsInfoEnabled) log.Info("Number of SEI users for updating - " + cnt);
            }

            finally
            {
                oConnection.Close();
            }

        }
    }
}
