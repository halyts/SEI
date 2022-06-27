-----------------------------------------------------Завантаження списку співробітників - учасників синхронізації.
-- namespace SEI
-- class Settings
-- public void ReloadSettings() 

select user_login, email_addr, direction_id, sync_start_date,SYNC_TOKEN, LAST_SYNC_DATE, USER_ID from cx_sei_users;

-----------------------------------------------------Завантаження списку видалених в Siebel  та несинхронізованих з Microsoft Exchange зустрічей.
--namespace SEI
--class SiebelDBW
--public static List<SiebelAppointment> GetDeleted()

select c.action_id, c.exch_id, c.exc_owner_appt_id
                                      from  cx_sei_commands c
                                     where c.command_executed = 'N'
                                       and c.command = 'DELETE_APPT'
                                       and c.exc_owner_appt_id is not null;

----------------------------------------------------UPDATE!!! Зміна статусів видалених в Siebel та синхронізованих з Microsoft Exchange зустрічей.

--namespace SEI
--class SiebelDBW
--public static void UpdateDeletedStatus(string Id)

/*
update cx_sei_commands c
                                         set c.command_executed = 'Y', last_upd = :2, last_upd_by = 'SEI', db_last_upd = :2, db_last_upd_src = 'SEI' where c.action_id = :1;
*/                                         

----------------------------------------------------Отримання дати із SIebel
--namespace SEI
--class SiebelDBW
--public static DateTime GetCurrentDT()

select sysdate from dual;

----------------------------------------------------Визначення Id технічного користувача в Siebel

--namespace SEI
--class SiebelDBW
--public static List<SiebelAppointment> LoadUserAppointments(SEIUser user)

select row_id from siebel.s_user where login = 'SEI_TECH_USER'; -- (login із налаштувань);

----------------------------------------------------Отримання списку зустрічей по користувачу доданих після останньої синхронізації

--namespace SEI
--class SiebelDBW
--public static List<SiebelAppointment> LoadUserAppointments(SEIUser user)
                select a.row_id,
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
                                            from s_evt_act a
                                            left join cx_sei_act_x sa
                                            on sa.par_row_id = a.row_id
                                            join s_contact c
                                            on c.row_id = a.owner_per_id
                                            join (select distinct row_id, when
                                                    from (select /*distinct*/
                                                            row_id,
                                                            max(last_upd_by) keep(dense_rank last order by last_upd) over(partition by row_id) who,
                                                            max(last_upd) over(partition by row_id) when
                                                            from (select a.row_id, a.last_upd, a.last_upd_by
                                                                    from s_evt_act a
                                                                    join s_lst_of_val lv on lv.type ='ACTIVITY_DISPLAY_CODE' and lv.name = 'Calendar and Activity' and lv.lang_id = 'RUS' and lv.val = a.cal_type_cd
                                                                    left join cx_sei_act_x sa
                                                                    on sa.par_row_id = a.row_id
                                                                    where a.OWNER_PER_ID = :2
                                                                    /*and a.cal_type_cd =
                                                                        (select val from s_lst_of_val where type ='ACTIVITY_DISPLAY_CODE' and name = 'Calendar and Activity' and lang_id = 'RUS')*/
                                                                    and (sa.sync_status is not null and
                                                                        sa.sync_status not in ('ED', 'CR','ExEtS'))
                                                                union all
                                                                select a.row_id, ae.last_upd, ae.last_upd_by
                                                                    from s_evt_act a
                                                                    join s_lst_of_val lv
                                                                    on lv.type ='ACTIVITY_DISPLAY_CODE' and lv.name = 'Calendar and Activity' and lv.lang_id = 'RUS' and lv.val = a.cal_type_cd
                                                                    left join cx_sei_act_x sa
                                                                    on sa.par_row_id = a.row_id
                                                                    join s_act_emp ae
                                                                    on ae.activity_id = a.row_id
                                                                    where a.OWNER_PER_ID = :2
                                                                    and (sa.sync_status is not null and
                                                                        sa.sync_status not in ('ED', 'CR','ExEtS'))
                                                                union all
                                                                select a.row_id, ac.last_upd, ac.last_upd_by
                                                                    from s_evt_act a
                                                                    join s_lst_of_val lv
                                                                    on lv.type ='ACTIVITY_DISPLAY_CODE' and lv.name = 'Calendar and Activity' and lv.lang_id = 'RUS' and lv.val = a.cal_type_cd
                                                                    left join cx_sei_act_x sa
                                                                    on sa.par_row_id = a.row_id
                                                                    join s_act_contact ac
                                                                    on ac.activity_id = a.row_id
                                                                    where a.OWNER_PER_ID = :2
                                                                    and (sa.sync_status is not null and
                                                                        sa.sync_status not in ('ED', 'CR','ExEtS'))
                                                                union all
                                                                select a.row_id, ar.last_upd, ar.last_upd_by
                                                                    from s_evt_act a
                                                                    join s_lst_of_val lv
                                                                    on lv.type ='ACTIVITY_DISPLAY_CODE' and lv.name = 'Calendar and Activity' and lv.lang_id = 'RUS' and lv.val = a.cal_type_cd
                                                                    left join cx_sei_act_x sa
                                                                    on sa.par_row_id = a.row_id
                                                                    join S_ACT_CAL_RSRC ar
                                                                    on ar.activity_id = a.row_id
                                                                    where a.OWNER_PER_ID = :2
                                                                    and (sa.sync_status is not null and
                                                                        sa.sync_status not in ('ED', 'CR','ExEtS'))
                                                                union all
                                                                select a.row_id, sc.last_upd, sc.last_upd_by
                                                                    from s_evt_act a
                                                                    join s_lst_of_val lv
                                                                    on lv.type ='ACTIVITY_DISPLAY_CODE' and lv.name = 'Calendar and Activity' and lv.lang_id = 'RUS' and lv.val = a.cal_type_cd
                                                                    left join cx_sei_act_x sa
                                                                    on sa.par_row_id = a.row_id
                                                                    join cx_sei_commands sc
                                                                    on sc.command in
                                                                        ('DELETE_EMP', 'DELETE_CON', 'DELETE_RES')
                                                                    and sc.action_id = a.row_id
                                                                    where a.OWNER_PER_ID = :2
                                                                    and (sa.sync_status is not null and
                                                                        sa.sync_status not in ('ED', 'CR', 'ExEtS'))))
                                                    where who != :1
                                                    and when >= :3) x
                                            on x.row_id = a.row_id;
                                            
 --Параметри
 --                                           1: 1-2NKXVVY                             Ідентифікатор в Siebel технічного користувача 
 --                                           2: 1-3AWSFIQ                             Ідентифікатор в Siebe користувача, що синхронізується
 --                                           3: 28.10.2019 15:41:04              Дата останньої синхронізації користувача, що синхронізується.


        
------------------------------------------------------Отримання списку обов'язкових учасників-працівників  для зустрічі
--namespace SEI
--class SiebelDBW
--public static List<SiebelAppointment> LoadUserAppointments(SEIUser user)

SELECT DISTINCT
    e.emp_id,
    c.last_name
    || ' '
    || c.fst_name
    ||
        CASE
            WHEN c.mid_name IS NOT NULL THEN
                ' ' || c.mid_name
            ELSE
                NULL
        END,
    c.email_addr,
    e.act_invt_resp_cd
FROM
    s_act_emp   e
    JOIN s_contact   c ON c.row_id = e.emp_id
WHERE
    e.activity_id = :1
    AND e.emp_id != :2
    AND e.x_sei_is_required = 'Y'
    AND c.email_addr IS NOT NULL;
--Параметри
--                   1: 1-474GKIH                          Ідентифікатор взаємодії в Siebel;
--                   2: 1-3AWSFIQ                        Ідентифікатор працівника, що створив зустріч;
                    



----------------------------------------------------------Отримання списку обов'язкових учасників - контактів для зустрічі (з наявним "Основной Email")
--namespace SEI
--class SiebelDBW
--public static List<SiebelAppointment> LoadUserAppointments(SEIUser user)


SELECT DISTINCT
    x.con_id,
    c.last_name
    || ' '
    || c.fst_name
    ||
        CASE
            WHEN c.mid_name IS NOT NULL THEN
                ' ' || c.mid_name
            ELSE
                NULL
        END,
    cd.value
FROM
    s_act_contact     x
    JOIN s_contact         c ON c.row_id = x.con_id
    JOIN cx_contact_data   cd ON cd.par_row_id = c.row_id
                               AND cd.active = 'Y'
                               AND cd.contact_type = 'Основной Email'
WHERE
    x.activity_id = :1
    AND x.con_id != :2
    AND x.x_sei_is_required = 'Y'
    AND cd.value IS NOT NULL;


--                   1: 1-474GKIH                          Ідентифікатор взаємодії в Siebel;
--                   2: 1-3AWSFIQ                        Ідентифікатор працівника, що створив зустріч;

------------------------------------------------------Отримання списку необов'язкових учасників-працівників  для зустрічі
--namespace SEI
--class SiebelDBW
--public static List<SiebelAppointment> LoadUserAppointments(SEIUser user)

SELECT DISTINCT
    e.emp_id,
    c.last_name
    || ' '
    || c.fst_name
    ||
        CASE
            WHEN c.mid_name IS NOT NULL THEN
                ' ' || c.mid_name
            ELSE
                NULL
        END,
    c.email_addr,
    e.act_invt_resp_cd
FROM
    s_act_emp   e
    JOIN s_contact   c ON c.row_id = e.emp_id
WHERE
    e.activity_id = :1
    AND e.emp_id != :2
    AND ( e.x_sei_is_required = 'N'
          OR e.x_sei_is_required IS NULL )
    AND c.email_addr IS NOT NULL;
    
--                   1: 1-474GKIH                          Ідентифікатор взаємодії в Siebel;
--                   2: 1-3AWSFIQ                        Ідентифікатор працівника, що створив зустріч;
    
----------------------------------------------------------Отримання списку необов'язкових учасників - контактів для зустрічі (з наявним "Основной Email")    
--namespace SEI
--class SiebelDBW
--public static List<SiebelAppointment> LoadUserAppointments(SEIUser user)


SELECT DISTINCT
    x.con_id,
    c.last_name
    || ' '
    || c.fst_name
    ||
        CASE
            WHEN c.mid_name IS NOT NULL THEN
                ' ' || c.mid_name
            ELSE
                NULL
        END,
    cd.value
FROM
    s_act_contact     x
    JOIN s_contact         c ON c.row_id = x.con_id
    JOIN cx_contact_data   cd ON cd.par_row_id = c.row_id
                               AND cd.active = 'Y'
                               AND cd.contact_type = 'Основной Email'
WHERE
    x.activity_id = :1
    AND x.con_id != :2
    AND ( x.x_sei_is_required = 'N'
          OR x.x_sei_is_required IS NULL )
    AND cd.value IS NOT NULL;
    


--                   1: 1-474GKIH                          Ідентифікатор взаємодії в Siebel;
--                   2: 1-3AWSFIQ                        Ідентифікатор працівника, що створив зустріч;


----------------------------------------------------------Отримання списку ресурсів для зустрічі 
--namespace SEI
--class SiebelDBW
--public static List<SiebelAppointment> LoadUserAppointments(SEIUser user)


SELECT DISTINCT
    r.row_id,
    r.name,
    r.x_sei_res_addr
FROM
    s_act_cal_rsrc   ar
    JOIN s_cal_rsrc       r ON r.row_id = ar.resource_id
WHERE
    ar.activity_id = :1;
    

--                   1: 1-474GKIH                          Ідентифікатор взаємодії в Siebel;


----------------------------------------------------Визначення Id технічного користувача в Siebel

--namespace SEI
--class SiebelDBW
--public static List<SiebelAppointment>LoadUserResponses(SEIUser user)

SELECT
    row_id
FROM
    s_user
WHERE
    login = :1;
    
--        1:  SEI_TECH_USER                     Логін технічного користувача із налаштувань


-----------------------------------------------------Завантаження відповідей користувачів
--namespace SEI
--class SiebelDBW
--public static List<EAppointmentResponse> LoadUserResponses(SEIUser user)

SELECT
    sa.exc_id,
    ae.last_upd,
    ae.x_change_status_dt,
    ae.act_invt_resp_cd
FROM
    s_evt_act      a
    JOIN cx_sei_act_x   sa ON sa.par_row_id = a.row_id
    JOIN s_act_emp      ae ON ae.activity_id = a.row_id
WHERE
    ae.x_change_status_dt > :3
    AND ae.emp_id = :2
    AND sa.exc_id IS NOT NULL
    AND ae.x_change_status_dt IS NOT NULL;
    
--    2:  Ідентифікатор користувача в Siebel
--    3:  Дата останньої синхронізації


    