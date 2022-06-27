using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;
using System.Net;
using log4net;
using log4net.Appender;
using log4net.Layout;
 

namespace SEI
{
    class Syncee
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public static bool SyncSiebelDeleted()
        {
            using (ThreadContext.Stacks["NDC"].Push("Siebel deleting"))
            {

                bool ret = false;
                try
                {

                    //Logger.Log("SyncSiebelDeleted Start");

                    List<SiebelAppointment> siebelDeletedAppointments = SiebelDBW.GetDeleted();

                    if (log.IsInfoEnabled) log.Info(String.Format("{0} appointments for deleting", siebelDeletedAppointments.Count));
                    //Logger.Log(String.Format("{0} appointments to delete", siebelDeletedAppointments.Count));

                    foreach (SiebelAppointment a in siebelDeletedAppointments)
                    {
                        try
                        {
                            if (log.IsInfoEnabled) log.Info("Deleting appointment " + a.Id + "->" + a.OwnerId);
                            //Logger.Log("Deleting appointment " + a.Id + "->" + a.OwnerId);

                            ExchangeWSW.DeleteAppointment(a);
                            SiebelDBW.UpdateDeletedStatus(a.SiebelId);
                        }
                        catch (Exception e)
                        {
                            if (e is ServiceResponseException)
                                if (!(((ServiceResponseException)e).ErrorCode == ServiceError.ErrorItemNotFound))
                                    throw (e);
                                else
                                    SiebelDBW.UpdateDeletedStatus(a.SiebelId);
                            else
                                throw (e);
                        }
                    }
                    ret = true;
                }
                catch (Exception e1)
                {
                    log.Error(e1.Message);
                    //Logger.Log(e1.Message, Logger.LogSeverity.Error);
                }

                //Logger.Log("SyncSiebelDeleted End");
                return ret;
            }
        }

        public static bool LoadRooms(out List<String> rooms)
        {
            using (ThreadContext.Stacks["NDC"].Push("Loading rooms"))
            {

                bool ret = false;
                List<String> oRooms = new List<string>();
                try
                {
                    //Logger.Log("LoadRooms Start");
                    oRooms = DomainW.GetRooms();
                    if (log.IsInfoEnabled) log.Info("Loaded " + oRooms.Count + " rooms");
                    ret = true;
                }
                catch (Exception e1)
                {
                    log.Error(e1.Message);
                    //Logger.Log(e1.Message, Logger.LogSeverity.Error);
                }

                //Logger.Log("LoadRooms End");

                rooms = oRooms;
                return ret;
            }
        }


        public static List<EAppointmentResponse> GetResponses(SEIUser user, List<SEIUser> userlist, ExchangeAppointments exchangeAppointments)
        {
            List<EAppointmentResponse> exchangeResponses = new List<EAppointmentResponse>();
            List<EAppointmentResponse> siebelResponses = new List<EAppointmentResponse>();
            List<EAppointmentResponse> res = new List<EAppointmentResponse>();

            //loading responses

            if (log.IsInfoEnabled) log.Info("Seeking Exchange responses");
            //Logger.Log("Seeking Exchange responses");
 
            foreach (SEIItemChange sEIItemChange in exchangeAppointments.Changes.Where(f => f.Change.ChangeType == ChangeType.Create || f.Change.ChangeType == ChangeType.Update).Where(f1 => !((Appointment)f1.Change.Item).IsCancelled /*&& !((Appointment)f1.Change.Item).IsRecurring*/))
            {
                //организатор затягивает всех, кто не в списке синхронизации
                if (
                    ((Appointment)sEIItemChange.Change.Item).Organizer.Address.Equals(user.Email, StringComparison.InvariantCultureIgnoreCase)
                    )
                {
                    foreach (Attendee attendee in ((Appointment)sEIItemChange.Change.Item).RequiredAttendees.Where(f => f.LastResponseTime.HasValue
                        &&
                        !userlist.Any(f1 => f1.Email.Equals(f.Address, StringComparison.InvariantCultureIgnoreCase))
                    ))
                        exchangeResponses.Add(new EAppointmentResponse(sEIItemChange.CGOID, attendee.ResponseType.ToString(), attendee.LastResponseTime, attendee.Address, "Exchange"));
                    foreach (Attendee attendee in ((Appointment)sEIItemChange.Change.Item).OptionalAttendees.Where(f => f.LastResponseTime.HasValue
                        &&
                        !userlist.Any(f1 => f1.Email.Equals(f.Address, StringComparison.InvariantCultureIgnoreCase))
                    ))
                        exchangeResponses.Add(new EAppointmentResponse(sEIItemChange.CGOID, attendee.ResponseType.ToString(), attendee.LastResponseTime, attendee.Address, "Exchange"));
                }
                //участник затягивает только себя
                else //if (!userlist.Any(f => f.Email.Equals(((Appointment)sEIItemChange.Change.Item).Organizer.Address, StringComparison.InvariantCultureIgnoreCase)))
                {
                    if (((Appointment)sEIItemChange.Change.Item).MyResponseType.ToString() != "NoResponseReceived")
                    {
                        DateTime? dtART = null;
                        try { dtART = ((Appointment)sEIItemChange.Change.Item).AppointmentReplyTime; }
                        catch (ServiceObjectPropertyException e)
                        {
                            if (e.Message != "This property was requested, but it wasn't returned by the server.")
                                throw;
                        }
                        if (dtART.HasValue)
                            exchangeResponses.Add(new EAppointmentResponse(sEIItemChange.CGOID, ((Appointment)sEIItemChange.Change.Item).MyResponseType.ToString(), dtART, user.Email, "Exchange"));
                    }
                }
            }

            if (log.IsInfoEnabled) log.Info("Found " + exchangeResponses.Count() + " raw Exchange responses");
            //Logger.Log("Found " + exchangeResponses.Count() + " raw Exchange responses");
            
            if (log.IsInfoEnabled) log.Info("Seeking Siebel responses");
            //Logger.Log("Seeking Siebel responses");

            siebelResponses = SiebelDBW.LoadUserResponses(user);

            if (log.IsInfoEnabled) log.Info("Found " + siebelResponses.Count() + " raw Siebel responses");
            //Logger.Log("Found " + siebelResponses.Count() + " raw Siebel responses");

//            IEnumerable<IGrouping<string, EAppointmentResponse>> query = ;

            foreach (IGrouping<string, EAppointmentResponse> gr in exchangeResponses.Concat(siebelResponses).GroupBy(f => f.CGOID + f.AttendeeAddress))
            {
                res.Add(gr.OrderBy(f => f.LastResponse).Last());
            }

            return res;
        }

        public static void SyncUser(SEIUser user, List<SEIUser> userlist, List<String> roomlist)
        {
            using (ThreadContext.Stacks["NDC"].Push(user.Login))
            {

                try
                {

                    //Logger.Log(String.Format("SyncUser Start: {0}", user.ToString()));
                    switch (user.Direction)
                    {
                        case "Bi":
                            //load siebelappointments + newsyncdate
                            //Logger.Log("Loading siebel appointments");
                            DateTime syncDate = SiebelDBW.GetCurrentDT();
                            List<SiebelAppointment> siebelAppointments = SiebelDBW.LoadUserAppointments(user);
                            if (siebelAppointments.Count() > 0)
                                syncDate = siebelAppointments[0].QueryDT;
                            if (log.IsInfoEnabled) log.Info("Loaded " + siebelAppointments.Count.ToString() + " Siebel appointments");
                            //Logger.Log("Loaded " + siebelAppointments.Count.ToString() + " appointments");

                            //load exchangeappointments + synctoken
                            //Logger.Log("Loading exchange appointments");
                            ExchangeAppointments exchangeAppointments = ExchangeWSW.GetAppointments(user.Email, user.SyncToken, user.SyncStartDate);
                            if (log.IsInfoEnabled) log.Info("Loaded " + exchangeAppointments.Changes.Count.ToString() + " Exchange appointments");
                            //Logger.Log("Loaded " + exchangeAppointments.Changes.Count.ToString() + " appointments");

                            //Logger.Log("Loading responses");
                            List<EAppointmentResponse> responses = GetResponses(user, userlist, exchangeAppointments);
                            if (log.IsInfoEnabled) log.Info("Loaded " + responses.Count.ToString() + " Exchange responses");
                            //Logger.Log("Loaded " + responses.Count.ToString() + " responses");

                            //logic
                            //-встречи могут быть назначены не участником синхронизации
                            //-встречи могут быть назначены только эмплоем
                            //-встречи, у которыхорганизатор не в эмплоях зибеля - игнорируются
                            //-события exchange создания удаления изменения
                            //-события siebel изменения удаления
                            //
                            //1. загружаем из зибеля удаленные встречи (CGUID), ищем встречи по овнерид, если нашли - удаляем встречу, проставляем статус удаленных команд.
                            //отменено, по списку далее//1.1. Загружаем из зибеля удаленных пользователей, ищем встречи по глобалид, если нашли - проставляем оклоненность встречи, проставляем статус удаленных команд.
                            //2. для каждого пользователя, подлежащего синхронизации
                            //2.1. загружаем все встречи, измененные в зибеле не пользователем синхронизации, с даты последней синхронизации siebel -> exchange
                            //2.2. загружаем все изменения папки пользователя в ексчейндже
                            //2.3. таблица решения
                            /*
                            |siebel  |exchange|exc own|action                                        |
                             | -      |created | N     | owner not in sync?transfer to siebel:nothing |
                             | -      |changed | N     | owner not in sync?transfer to siebel:nothing |
                             | -      |deleted | N     | try to delete - must fail as that id is not stored, owner status can be synced only by owner |
                   | -      | -      | N     | - |
                   |changed |created | N     | not possible yet |
                   |changed |changed | N     | not possible yet |
                   |changed |deleted | N     | not possible yet |
                   |changed | -      | N     | not possible yet |
                             | -      |created | Y     | transfer to siebel |
                             | -      |changed | Y     | transfer to siebel |
                             | -      |deleted | Y     | delete from siebel by owner id |
                   | -      | -      | Y     | - |
                            |changed |created | Y     | check update time and decide sync direction |
                            |changed |changed | Y     | check update time and decide sync direction |
                             |changed |deleted | Y     | delete from siebel by owner id |
                             |changed | -      | Y     | create in siebel |
                            */
                            //2.4. удаляем из зибеля все удаленные в ексчейндже встречи
                            //2.4.1. удаляем из зибеля все отмененные владельцем и для владельца в ексчейндже встречи
                            //2.5. передаём в ексчейндж все встречи, измененные в зибеле, для которых пустой cgoid или cgoid не в списке изменений из ексчейнджа
                            //2.6. передаём в зибель все встречи, созданные или измененные в ексчейндже, у которых cgoid не в списке изменений из зибеля. Если Organizer в списке синхронизации то в зибель встреча передается только в потоке синхронизации органайзера. Если Organizer не в списке синхронизации то в зибель встреча передается только в потоке синхронизации пользователя, который первый в списке участников, которые в списке синхронизации, отсортированном по логину.
                            //2.7. для встреч, которые и там и там - определяем направление передачи основываясь на дате последнего изменения. Если Organizer в списке синхронизации то в зибель встреча передается только в потоке синхронизации органайзера. Если Organizer не в списке синхронизации то в зибель встреча передается только в потоке синхронизации пользователя, который первый в списке участников, которые в списке синхронизации, отсортированном по логину.


                            //2.4.
                            //Logger.Log("Deleting " + exchangeAppointments.Changes.Count(f => f.Change.ChangeType == ChangeType.Delete).ToString() + " exchange appointments from siebel");
                            foreach (SEIItemChange sEIItemChange in exchangeAppointments.Changes.Where(f => f.Change.ChangeType == ChangeType.Delete))
                            {
                                if (log.IsInfoEnabled) log.Info("Siebel deleting appointment " + sEIItemChange.Change.ItemId.UniqueId);
                                //Logger.Log("Deleting appointment " + sEIItemChange.Change.ItemId.UniqueId);
                                SiebelResult siebelResult = SiebelWSW.Delete(sEIItemChange.Change.ItemId.UniqueId, "");
                                if (log.IsInfoEnabled) log.Info(siebelResult.Status ? ("Siebel delete Success, " + siebelResult.Id) : "Siebel delete Fail");
                                //Logger.Log(siebelResult.Status ? ("Siebel delete Success, " + siebelResult.Id) : "Siebel delete Fail");
                                if (!siebelResult.Status)
                                {
                                    log.Error("Deleting Siebel appointment was unsuccessful");
                                    throw (new ApplicationException("Deleting appointment was unsuccessful"));
                                }
                            }

                            //2.4.1
                            foreach (SEIItemChange sEIItemChange in exchangeAppointments.Changes.Where(f => (f.Change.ChangeType == ChangeType.Create || f.Change.ChangeType == ChangeType.Update) && ((Appointment)f.Change.Item).IsCancelled))
                            {
                                if (log.IsInfoEnabled) log.Info("Canceling appointment " + sEIItemChange.Change.ItemId.UniqueId);
                                //Logger.Log("Canceling appointment " + sEIItemChange.Change.ItemId.UniqueId);
                                SiebelResult siebelResult = SiebelWSW.Delete(sEIItemChange.Change.ItemId.UniqueId, "");
                                if (log.IsInfoEnabled) log.Info(siebelResult.Status ? ("Siebel cancel Success, " + siebelResult.Id) : "Siebel cancel Fail");
                                //Logger.Log(siebelResult.Status ? ("Siebel cancel Success, " + siebelResult.Id) : "Siebel cancel Fail");
                                if (!siebelResult.Status)
                                {
                                    log.Error("Siebel canceling appointment was unsuccessful");
                                    throw (new ApplicationException("Canceling appointment was unsuccessful"));
                                }
                            }

                            //2.5.
                            //Logger.Log("Transfering " + siebelAppointments.Count(f => f.Id == "" || !exchangeAppointments.Changes.Where(f1 => f1.Change.ChangeType == ChangeType.Create || f1.Change.ChangeType == ChangeType.Update).Any(f2 => f2.CGOID == f.Id)).ToString() + " siebel origin appointments to exchange");
                            SEIId exchangeId;
                            foreach (SiebelAppointment a in siebelAppointments.Where(f => f.Id == "" || !exchangeAppointments.Changes.Where(f1 => f1.Change.ChangeType == ChangeType.Create || f1.Change.ChangeType == ChangeType.Update).Any(f2 => f2.CGOID == f.Id)))
                            {
                                if (a.Id == "")
                                {
                                    if (log.IsInfoEnabled) log.Info("Transfering appointment " + a.SiebelId + " to Exchange");
                                    //Logger.Log("Transfering appointment " + a.SiebelId + " to exchange");
                                    exchangeId = ExchangeWSW.CreateAppointment(a);
                                    if (exchangeId.CGOID == "" || exchangeId.LocalId == "")
                                    {
                                        SiebelDBW.SaveAppointmentSyncState(a, "ExStE");
                                        log.Error("Exchange create appointment was unsuccessful");
                                        throw (new ApplicationException("Exchange create appointment was unsuccessful"));
                                    }
                                    a.Id = exchangeId.CGOID;
                                    a.OwnerId = exchangeId.LocalId;
                                    if (log.IsInfoEnabled) log.Info("Saving appointment " + a.SiebelId + " as " + a.Id + ";" + a.OwnerId);
                                    //Logger.Log("Saving appointment " + a.SiebelId + " as " + a.Id + ";" + a.OwnerId);
                                    //save appointment excahnge id
                                    SiebelDBW.SaveAppointmentId(a);
                                    SiebelDBW.SaveAppointmentSyncState(a, "OK");
                                    SiebelWSW.ChangeActStatus(a.SiebelId, "Назначено");
                                    if (log.IsInfoEnabled) log.Info("Transfer of appointment " + a.SiebelId + " completed");
                                    //Logger.Log("Transfer of appointment " + a.SiebelId + " completed");
                                }
                                else
                                {
                                    if (log.IsInfoEnabled) log.Info("Transfering appointment " + a.SiebelId + " to Exchange");
                                    //Logger.Log("Transfering appointment " + a.SiebelId + " to exchange");
                                    exchangeId = ExchangeWSW.UpdateAppointment(a);
                                    if (exchangeId.CGOID == "" && exchangeId.LocalId == "")
                                    {
                                        if (log.IsInfoEnabled) log.Info("Deleting appointment " + a.SiebelId);
                                        //Logger.Log("Deleting appointment " + a.SiebelId);
                                        SiebelResult siebelResult = SiebelWSW.Delete("", a.SiebelId);
                                        if (log.IsInfoEnabled) log.Info(siebelResult.Status ? ("Siebel delete Success, " + siebelResult.Id) : "Siebel delete Fail");
                                        //Logger.Log(siebelResult.Status ? ("Siebel delete Success, " + siebelResult.Id) : "Siebel delete Fail");
                                        if (!siebelResult.Status)
                                        {
                                            log.Error("Deleting appointment was unsuccessful");
                                            throw (new ApplicationException("Deleting appointment was unsuccessful"));
                                        }
                                    }
                                    else if (exchangeId.CGOID == "" || exchangeId.LocalId == "")
                                    {
                                        SiebelDBW.SaveAppointmentSyncState(a, "ExStE");
                                        log.Error("Exchange update appointment was unsuccessful");
                                        throw (new ApplicationException("Exchange update appointment was unsuccessful"));
                                    }
                                    SiebelDBW.SaveAppointmentSyncState(a, "OK");
                                    SiebelWSW.ChangeActStatus(a.SiebelId, "Назначено");
                                    if (log.IsInfoEnabled) log.Info("Transfer of appointment " + a.SiebelId + " completed");
                                    //Logger.Log("Transfer of appointment " + a.SiebelId + " completed");
                                }
                            }

                            //2.6.
                            //Logger.Log("Transfering " + exchangeAppointments.Changes.Where(f => f.Change.ChangeType == ChangeType.Create || f.Change.ChangeType == ChangeType.Update).Count(f1 => !((Appointment)f1.Change.Item).IsCancelled && !siebelAppointments.Any(f2 => f2.Id == f1.CGOID)).ToString() + " exchange origin appointments to siebel");
                            foreach (SEIItemChange sEIItemChange in exchangeAppointments.Changes.Where(f => f.Change.ChangeType == ChangeType.Create || f.Change.ChangeType == ChangeType.Update).Where(f1 => !((Appointment)f1.Change.Item).IsCancelled && !siebelAppointments.Any(f2 => f2.Id == f1.CGOID)))
                            {

                                if (
                                    ((Appointment)sEIItemChange.Change.Item).Organizer.Address.Equals(user.Email, StringComparison.InvariantCultureIgnoreCase)
                                    )
                                {
                                    if (log.IsInfoEnabled) log.Info("Transferring appointment " + sEIItemChange.Change.ItemId.UniqueId + " to siebel");
                                    //Logger.Log("Transferring appointment " + sEIItemChange.Change.ItemId.UniqueId + " to siebel");
                                    SiebelResult siebelResult = SiebelWSW.InsertNew((Appointment)sEIItemChange.Change.Item, user, roomlist);
                                    if (siebelResult.Status)
                                    {
                                        if (log.IsInfoEnabled) log.Info("Siebel inserted success, " + siebelResult.Id);
                                        //Logger.Log("Siebel upsert Success, " + siebelResult.Id);
                                    }
                                    else
                                    {
                                        log.Error("Siebel upsert Fail");
                                        throw (new ApplicationException("Siebel upsert Fail"));
                                    }
                                }
                                else if (((Appointment)sEIItemChange.Change.Item).RequiredAttendees.Concat(((Appointment)sEIItemChange.Change.Item).OptionalAttendees).Where(f => userlist.Any(f1 => f1.Email.Equals(f.Address, StringComparison.InvariantCultureIgnoreCase))).OrderBy(f => f.Address).Count() > 0)
                                {
                                    if (
                                            //organizer not in synclist
                                            !userlist.Any(f => f.Email.Equals(((Appointment)sEIItemChange.Change.Item).Organizer.Address, StringComparison.InvariantCultureIgnoreCase))
                                            &&
                                            //current user - first in synced attendees
                                            user.Email.Equals(((Appointment)sEIItemChange.Change.Item).RequiredAttendees.Concat(((Appointment)sEIItemChange.Change.Item).OptionalAttendees).Where(f => userlist.Any(f1 => f1.Email.Equals(f.Address, StringComparison.InvariantCultureIgnoreCase))).OrderBy(f => f.Address).First().Address, StringComparison.InvariantCultureIgnoreCase)
                                        )
                                    {

                                        if (log.IsInfoEnabled) log.Info("Transferring appointment " + sEIItemChange.Change.ItemId.UniqueId + " to siebel");
                                        //Logger.Log("Transferring appointment " + sEIItemChange.Change.ItemId.UniqueId + " to siebel");
                                        SiebelResult siebelResult = SiebelWSW.InsertNew((Appointment)sEIItemChange.Change.Item, user, roomlist);
                                        if (siebelResult.Status)
                                        {
                                            if (log.IsInfoEnabled) log.Info("Siebel inserted success, " + siebelResult.Id);
                                            //Logger.Log("Siebel upsert Success, " + siebelResult.Id);
                                        }
                                        else
                                        {
                                            log.Error("Siebel upsert Fail");
                                            throw (new ApplicationException("Siebel upsert Fail"));
                                        }
                                    }
                                }
                                else
                                {
                                    if (log.IsInfoEnabled) log.Info("Skipping: " + sEIItemChange.Change.ItemId.UniqueId);
                                    //Logger.Log("Skipping: " + sEIItemChange.Change.ItemId.UniqueId);
                                }
                            }

                            //2.7.
                            //Logger.Log("Merging " + exchangeAppointments.Changes.Where(f => f.Change.ChangeType == ChangeType.Create || f.Change.ChangeType == ChangeType.Update).Count(f1 => !((Appointment)f1.Change.Item).IsCancelled && siebelAppointments.Any(f2 => f2.Id == f1.CGOID)).ToString() + " appointments");
                            foreach (SEIItemChange sEIItemChange in exchangeAppointments.Changes.Where(f => f.Change.ChangeType == ChangeType.Create || f.Change.ChangeType == ChangeType.Update).Where(f1 => !((Appointment)f1.Change.Item).IsCancelled && siebelAppointments.Any(f2 => f2.Id == f1.CGOID)))
                            {
                                if (sEIItemChange.Change.Item.LastModifiedTime >= siebelAppointments.Find(f => f.Id == sEIItemChange.CGOID).LastUpdated)
                                {
                                    if (
                                        ((Appointment)sEIItemChange.Change.Item).Organizer.Address.Equals(user.Email, StringComparison.InvariantCultureIgnoreCase)
                                        )
                                    {

                                        if (log.IsInfoEnabled) log.Info("Transferring appointment " + sEIItemChange.CGOID + " to siebel");
                                        //Logger.Log("Transferring appointment " + sEIItemChange.CGOID + " to siebel");
                                        SiebelResult siebelResult = SiebelWSW.InsertNew((Appointment)sEIItemChange.Change.Item, user, roomlist);
                                        if (siebelResult.Status)
                                        {
                                            if (log.IsInfoEnabled) log.Info("Siebel inserted success, " + siebelResult.Id);
                                            //Logger.Log("Siebel upsert Success, " + siebelResult.Id);
                                        }
                                        else
                                        {
                                            log.Error("Siebel upsert Fail");
                                            throw (new ApplicationException("Siebel upsert Fail"));
                                        }
                                    }
                                    else if (((Appointment)sEIItemChange.Change.Item).RequiredAttendees.Concat(((Appointment)sEIItemChange.Change.Item).OptionalAttendees).Where(f => userlist.Any(f1 => f1.Email.Equals(f.Address, StringComparison.InvariantCultureIgnoreCase))).OrderBy(f => f.Address).Count() > 0)
                                    {
                                        if (
                                                //organizer not in synclist
                                                !userlist.Any(f => f.Email.Equals(((Appointment)sEIItemChange.Change.Item).Organizer.Address, StringComparison.InvariantCultureIgnoreCase))
                                                &&
                                                //current user - first in synced attendees
                                                user.Email.Equals(((Appointment)sEIItemChange.Change.Item).RequiredAttendees.Concat(((Appointment)sEIItemChange.Change.Item).OptionalAttendees).Where(f => userlist.Any(f1 => f1.Email.Equals(f.Address, StringComparison.InvariantCultureIgnoreCase))).OrderBy(f => f.Address).First().Address, StringComparison.InvariantCultureIgnoreCase)
                                            )
                                        {
                                            if (log.IsInfoEnabled) log.Info("Transferring appointment " + sEIItemChange.CGOID + " to siebel");
                                            SiebelResult siebelResult = SiebelWSW.InsertNew((Appointment)sEIItemChange.Change.Item, user, roomlist);
                                            if (siebelResult.Status)
                                            {
                                                if (log.IsInfoEnabled) log.Info("Siebel inserted success, " + siebelResult.Id);
                                            }
                                            else
                                            {
                                                log.Error("Siebel insert fail");
                                                throw (new ApplicationException("Siebel upsert Fail"));
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (log.IsInfoEnabled) log.Info("Skipping - not owner: " + sEIItemChange.CGOID);
                                        //Logger.Log("Skipping - not owner: " + sEIItemChange.CGOID);
                                    }
                                }
                                else
                                {
                                    if (log.IsInfoEnabled) log.Info("Transfering appointment " + siebelAppointments.Find(f => f.Id == sEIItemChange.CGOID).SiebelId + " to exchange");
                                    //Logger.Log("Transfering appointment " + siebelAppointments.Find(f => f.Id == sEIItemChange.CGOID).SiebelId + " to exchange");
                                    exchangeId = ExchangeWSW.UpdateAppointment(siebelAppointments.Find(f => f.Id == sEIItemChange.CGOID));
                                    if (exchangeId.CGOID == "" || exchangeId.LocalId == "")
                                    {
                                        SiebelDBW.SaveAppointmentSyncState(siebelAppointments.Find(f => f.Id == sEIItemChange.CGOID), "ExStE");
                                        throw (new ApplicationException("Exchange update appointment was unsuccessful"));
                                    }
                                    SiebelDBW.SaveAppointmentSyncState(siebelAppointments.Find(f => f.Id == sEIItemChange.CGOID), "OK");
                                    if (log.IsInfoEnabled) log.Info("Transfer of appointment " + siebelAppointments.Find(f => f.Id == sEIItemChange.CGOID).SiebelId + " completed");
                                    //Logger.Log("Transfer of appointment " + siebelAppointments.Find(f => f.Id == sEIItemChange.CGOID).SiebelId + " completed");
                                }
                            }
                            if (log.IsInfoEnabled) log.Info("Processing responses");
                                //Logger.Log("Processing responses");
                                foreach (EAppointmentResponse response in responses)
                            {
                                if (response.Source == "Siebel")
                                    ExchangeWSW.ResponseToAppointment(response.CGOID, response.AttendeeAddress, response.Response);
                                if (response.Source == "Exchange")
                                    SiebelDBW.SetResponse(response.CGOID, syncDate, response.AttendeeAddress, response.Response);
                            }
                            if (log.IsInfoEnabled) log.Info(responses.Count + " responses completed");

                            //if (log.IsDebugEnabled) log.Debug("Saving lastsyncdate " + (siebelAppointments.Count > 0 ? siebelAppointments[0].QueryDT.ToString() : "-") + " and synctoken " + exchangeAppointments.SyncToken);
                            //save lastsyncdate and synctoken

                            //My ---------------------------------------------------------------------------------------------------------------------------------

                            exchangeAppointments = ExchangeWSW.GetAppointments(user.Email, user.SyncToken, user.SyncStartDate);
                            //------------------------------------------------------------------------------------------------------------------------------------
                            SiebelDBW.SaveUserSyncState(user.Login, syncDate, exchangeAppointments.SyncToken);
                            if (log.IsInfoEnabled) log.Info("User synchronization completed");
                            //Logger.Log("User sync completed");
                            break;
                        default:
                            if (log.IsInfoEnabled) log.Info("Direction: " + user.Direction + " not supported yet");
                            //Logger.Log("Direction: " + user.Direction + " not supported yet");
                            break;
                    }
                }
                catch (Exception e)
                {
                    log.Error(e.Message);
                    //Logger.Log(e.Message, Logger.LogSeverity.Error);
                    SiebelDBW.SaveUserError(user.Login, DateTime.Now.ToString() + "; " + e.Message);
                }
                //catch - try save lasterror
                if (log.IsInfoEnabled) log.Info("User synchronization end");
                //Logger.Log("SyncUser End");
            }
        }
    }
}
