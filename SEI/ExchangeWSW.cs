using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;

using log4net;

namespace SEI
{
    class ExchangeAppointments
    {
        public List<SEIItemChange> Changes { get; set; }
        public string SyncToken { get; set; }
        public ExchangeAppointments ()
        {
            this.Changes = new List<SEIItemChange>();
            this.SyncToken = "";
        }
    }
    class EAppointmentResponse
    {
        public string CGOID { get; set; }
        public string Response { get; set; }
        public DateTime? LastResponse { get; set; }
        public string AttendeeAddress { get; set; }
        public string Source { get; set; }
        public EAppointmentResponse(string cGOID, string response, DateTime? lastResponse, string attendeeAddress, string source)
        {
            this.CGOID = cGOID;
            this.Response = response;
            this.LastResponse = lastResponse;
            this.AttendeeAddress = attendeeAddress;
            this.Source = source;
        }
    }
    class SEIItemChange 
    {
        public ItemChange Change { get; set; }
        public string CGOID { get; set; }
        public SEIItemChange (ItemChange change, string cGOID)
        {
            this.Change = change;
            this.CGOID = cGOID;
        }
    }
    class SEIId
    {
        public string CGOID { get; set; }
        public string LocalId { get; set; }
        public SEIId (string cGOID, string localId)
        {
            this.CGOID = cGOID;
            this.LocalId = localId;
        }
    }

    class ExchangeWSW
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        
        public static void DeleteAppointment(SiebelAppointment a)
        {
            ExchangeService exService = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
            exService.Timeout = Properties.Settings.Default.ExchangeWSTimeout * 1000;
            exService.Url = new Uri(Properties.Settings.Default.ExchangeService);
            exService.Credentials = new NetworkCredential(Properties.Settings.Default.ExchangeLogin, Properties.Settings.Default.uExchangePassword, Properties.Settings.Default.ExchangeDomain);
            exService.TraceListener = new EWSTraceListener();
            exService.TraceFlags = TraceFlags.EwsRequest | TraceFlags.EwsResponse;
            exService.TraceEnabled = true;
            Appointment appt = Appointment.Bind(exService, new ItemId(a.OwnerId));
            appt.Delete(DeleteMode.MoveToDeletedItems, SendCancellationsMode.SendOnlyToAll);
        }

        public static void ResponseToAppointment(string CGOID, string email, string response)
        {
            using (ThreadContext.Stacks["NDC"].Push("Find Exchange items"))
            {

                ExtendedPropertyDefinition MAPICleanGlobalObjectId = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.Meeting, 0x23, MapiPropertyType.Binary);

                ExchangeService exService = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
                exService.Timeout = Properties.Settings.Default.ExchangeWSTimeout * 1000;
                exService.Url = new Uri(Properties.Settings.Default.ExchangeService);
                exService.Credentials = new NetworkCredential(Properties.Settings.Default.ExchangeLogin, Properties.Settings.Default.uExchangePassword, Properties.Settings.Default.ExchangeDomain);
                exService.TraceListener = new EWSTraceListener();
                exService.TraceFlags = TraceFlags.EwsRequest | TraceFlags.EwsResponse;
                exService.TraceEnabled = true;
                CalendarFolder calendar = CalendarFolder.Bind(exService, new FolderId(WellKnownFolderName.Calendar, email));
                FindItemsResults<Item> foundItems = calendar.FindItems(new SearchFilter.IsEqualTo(MAPICleanGlobalObjectId, CGOID), new ItemView(1));
                ItemId aid = null;
                if (foundItems.Items.Count > 0)
                    aid = foundItems.Items[0].Id;
                if (aid == null)
                {
                    if (log.IsInfoEnabled) log.Info("Item not found in Calendar, searching in Deleted");
                    //Logger.Log("Item not found in Calendar, searching in Deleted");
                    Folder deleted = Folder.Bind(exService, new FolderId(WellKnownFolderName.DeletedItems, email));
                    foundItems = deleted.FindItems(new SearchFilter.IsEqualTo(MAPICleanGlobalObjectId, CGOID), new ItemView(1));
                    if (foundItems.Items.Count > 0)
                        aid = foundItems.Items[0].Id;
                }
                if (aid == null)
                {
                    if (log.IsWarnEnabled) log.Warn("Item not found in Deleted, searching in all folders");
                    //Logger.Log("Item not found in Deleted, searching in all folders", Logger.LogSeverity.Warning);
                    FindFoldersResults folders = exService.FindFolders(new FolderId(WellKnownFolderName.Root, email), new FolderView(1000));
                    foreach (Folder folder in folders.Where(f => !f.GetType().Equals(typeof(SearchFolder))))
                    {
                        if (folder.Id != (new FolderId(WellKnownFolderName.DeletedItems, email))
                            && folder.Id != (new FolderId(WellKnownFolderName.Calendar, email))
                           )
                        {
                            try
                            {
                                foundItems = folder.FindItems(new SearchFilter.IsEqualTo(MAPICleanGlobalObjectId, CGOID), new ItemView(1));
                                if (foundItems.Items.Count > 0)
                                {
                                    if (log.IsWarnEnabled) log.Warn("Item found in " + folder.DisplayName);
                                    //Logger.Log("Item found in " + folder.DisplayName, Logger.LogSeverity.Warning);
                                    aid = foundItems.Items[0].Id;
                                    //break;
                                }
                            }
                            catch (ServiceResponseException e)
                            {
                                if (e.ErrorCode != ServiceError.ErrorAccessDenied)
                                    throw;
                            }
                        }
                    }
                }

                if (aid != null)
                {
                    if (foundItems.Items[0].GetType() == typeof(Appointment))
                    {
                        Appointment apt = Appointment.Bind(exService,
                        aid,
                        new PropertySet(
                            BasePropertySet.FirstClassProperties
                            )
                        );
                        switch (response)
                        {
                            case "Tentative":
                                //if (apt.MyResponseType != MeetingResponseType.Tentative)
                                if (!apt.IsCancelled)
                                    apt.AcceptTentatively(true);
                                break;
                            case "Accepted":
                                //if(apt.MyResponseType != MeetingResponseType.Accept)
                                if (!apt.IsCancelled)
                                    apt.Accept(true);
                                break;
                            case "Declined":
                                //if (apt.MyResponseType != MeetingResponseType.Decline)
                                if (!apt.IsCancelled)
                                    apt.Decline(true);
                                break;
                            default:
                                break;
                        }
                    }
                    else if (foundItems.Items[0].GetType() == typeof(MeetingRequest))
                    {
                        MeetingRequest mr = MeetingRequest.Bind(exService, aid, new PropertySet(
                            BasePropertySet.FirstClassProperties
                            ));
                        switch (response)
                        {
                            case "Tentative":
                                //if (apt.MyResponseType != MeetingResponseType.Tentative)
                                if (!mr.IsCancelled)
                                    mr.AcceptTentatively(true);
                                break;
                            case "Accepted":
                                //if(apt.MyResponseType != MeetingResponseType.Accept)
                                if (!mr.IsCancelled)
                                    mr.Accept(true);
                                break;
                            case "Declined":
                                //if (apt.MyResponseType != MeetingResponseType.Decline)
                                if (!mr.IsCancelled)
                                    mr.Decline(true);
                                break;
                            default:
                                break;
                        }
                    }
                    else
                        throw new ApplicationException("Unknown type of object to response: " + foundItems.Items[0].GetType().Name);
                }
                else
                {
                    if (log.IsWarnEnabled) log.Warn("Item not found");
                    //Logger.Log("Item not found", Logger.LogSeverity.Warning);
                }
            }
        }

        public static ExchangeAppointments GetAppointments(string UserEmail, string SyncToken, DateTime SyncFrom)
        {
            ExchangeAppointments res = new ExchangeAppointments();
            ExtendedPropertyDefinition CleanGlobalObjectId = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.Meeting, 0x23, MapiPropertyType.Binary);
            object CGOID;

            ExchangeService exService = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
            exService.Timeout = Properties.Settings.Default.ExchangeWSTimeout * 1000;
            exService.Url = new Uri(Properties.Settings.Default.ExchangeService);
            exService.Credentials = new NetworkCredential(Properties.Settings.Default.ExchangeLogin, Properties.Settings.Default.uExchangePassword, Properties.Settings.Default.ExchangeDomain);
            exService.TraceListener = new EWSTraceListener();
            exService.TraceFlags = TraceFlags.EwsRequest | TraceFlags.EwsResponse;
            exService.TraceEnabled = true;
            CalendarFolder calendar = CalendarFolder.Bind(exService, new FolderId(WellKnownFolderName.Calendar, UserEmail));
            ChangeCollection<ItemChange> changes = exService.SyncFolderItems(calendar.Id, BasePropertySet.FirstClassProperties, null, 10, SyncFolderItemsScope.NormalItems, SyncToken);
            //if (log.IsDebugEnabled) log.Debug("Changes: " + changes.Count);
            //res.Changes.AddRange(changes.Where(f => f.Item.GetType() == typeof(Appointment))); 4 delete item = null
            foreach (ItemChange ic in changes)
            {
                if (ic.ChangeType == ChangeType.Create || ic.ChangeType == ChangeType.Update)
                {
                    if ((SyncToken == "" && (ic.Item.DateTimeCreated >= SyncFrom || ((Appointment)ic.Item).Start >= SyncFrom)) || SyncToken != "")
                    {
                        ic.Item.Load(
                            new PropertySet(
                                BasePropertySet.FirstClassProperties,
                                AppointmentSchema.Body,
                                AppointmentSchema.AdjacentMeetingCount,
                                AppointmentSchema.MimeContent,
                                AppointmentSchema.UniqueBody,
                                AppointmentSchema.Recurrence,
                                AppointmentSchema.RequiredAttendees,
                                CleanGlobalObjectId
                                )
                            { RequestedBodyType = BodyType.Text }
                            );
                        ic.Item.TryGetProperty(CleanGlobalObjectId, out CGOID);
                        res.Changes.Add(new SEIItemChange(ic, Convert.ToBase64String((Byte[])CGOID)));
                    }
                }
                else if(ic.ChangeType == ChangeType.Delete)
                {
                    res.Changes.Add(new SEIItemChange(ic, ""));
                }
            }
            res.SyncToken = changes.SyncState;
            while (changes.MoreChangesAvailable)
            {
                changes = exService.SyncFolderItems(calendar.Id, BasePropertySet.FirstClassProperties, null, 10, SyncFolderItemsScope.NormalItems, res.SyncToken);

                foreach (ItemChange ic in changes)
                {
                    if (ic.ChangeType == ChangeType.Create || ic.ChangeType == ChangeType.Update)
                    {
                        if ((SyncToken == "" && (ic.Item.DateTimeCreated >= SyncFrom || ((Appointment)ic.Item).Start >= SyncFrom)) || SyncToken != "")
                        {
                            ic.Item.Load(
                                new PropertySet(
                                    BasePropertySet.FirstClassProperties,
                                    AppointmentSchema.Body,
                                    AppointmentSchema.AdjacentMeetingCount,
                                    AppointmentSchema.MimeContent,
                                    AppointmentSchema.UniqueBody,
                                    AppointmentSchema.Recurrence,
                                    CleanGlobalObjectId
                                    )
                                { RequestedBodyType = BodyType.Text }
                                );
                            ic.Item.TryGetProperty(CleanGlobalObjectId, out CGOID);
                            res.Changes.Add(new SEIItemChange(ic, Convert.ToBase64String((Byte[])CGOID)));
                        }
                    }
                    else if (ic.ChangeType == ChangeType.Delete)
                    {
                        res.Changes.Add(new SEIItemChange(ic, ""));
                    }
                }

                res.SyncToken = changes.SyncState;
            }
            return res;
        }

        public static SEIId CreateAppointment(SiebelAppointment appointment)
        {
            ExtendedPropertyDefinition MAPICleanGlobalObjectId = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.Meeting, 0x23, MapiPropertyType.Binary);
            ExtendedPropertyDefinition MAPILocationDisplayName = new ExtendedPropertyDefinition(new Guid("A719E259-2A9A-4FB8-BAB3-3A9F02970E4B"), "LocationDisplayName", MapiPropertyType.String);
            ExtendedPropertyDefinition MAPILocationUri = new ExtendedPropertyDefinition(new Guid("A719E259-2A9A-4FB8-BAB3-3A9F02970E4B"), "LocationUri", MapiPropertyType.String);
            ExtendedPropertyDefinition MAPILocationSource = new ExtendedPropertyDefinition(new Guid("A719E259-2A9A-4FB8-BAB3-3A9F02970E4B"), "LocationSource", MapiPropertyType.Integer);
            ExtendedPropertyDefinition MAPILocationAnnotation = new ExtendedPropertyDefinition(new Guid("A719E259-2A9A-4FB8-BAB3-3A9F02970E4B"), "LocationAnnotation", MapiPropertyType.String);
            object oMAPICleanGlobalObjectId;
            //object oMAPILocationDisplayName;
            //object oMAPILocationAnnotation;
            //object oMAPILocationSource;
            //object oMAPILocationUri;

            ExchangeService exService = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
            exService.Timeout = Properties.Settings.Default.ExchangeWSTimeout * 1000;
            exService.Url = new Uri(Properties.Settings.Default.ExchangeService);
            exService.Credentials = new NetworkCredential(Properties.Settings.Default.ExchangeLogin, Properties.Settings.Default.uExchangePassword, Properties.Settings.Default.ExchangeDomain);
            exService.TraceListener = new EWSTraceListener();
            exService.TraceFlags = TraceFlags.EwsRequest | TraceFlags.EwsResponse;
            exService.TraceEnabled = true;
            CalendarFolder calendar = CalendarFolder.Bind(exService, new FolderId(WellKnownFolderName.Calendar, appointment.Owner.Email));

            Appointment a = new Appointment(exService);
            string clientName = string.Empty;
            string changedSubject = string.Empty;
            string respName = string.Empty;
            int Offset = appointment.Subject.IndexOf("Ответственный:");
            if (Offset != -1)
            {
                string exchDomain = Properties.Settings.Default.ExchangeDomain;
                int OffsetClient = appointment.Subject.IndexOf("Клиент:");
                clientName = appointment.Subject.Substring(OffsetClient + 7, Offset - OffsetClient - 7);
                changedSubject = appointment.Subject.Remove(OffsetClient);
                respName = appointment.Subject.Substring(Offset + 14);
                MessageBody mesBody = new MessageBody(BodyType.HTML, "Добрый день!<br>Назначена встреча с клиентом: " + clientName + "<br>Ответственный:" + respName +
                                     "<br>Место встречи: " + appointment.MeetingLocation + "<br>Дата и время встречи: " + 
                                     appointment.StartDate.ToString("dd.MM.yyyy HH:mm") + "<br>Комментарий: Необходимо подтвердить дату и время встречи с клиентом за день<br>Ваш Сбербанк.");
                a.Subject = changedSubject;
                a.Body = mesBody;                               
            }
            else
            {
                a.Subject = appointment.Subject;
                a.Body = appointment.Body;
            }

            a.Start = appointment.StartDate;
            a.End = appointment.EndDate;
            a.ReminderMinutesBeforeStart = appointment.ReminderMinutesBeforeStart;

            foreach (SiebelAttendee attendee in appointment.RequiredAttendees)
            {
                a.RequiredAttendees.Add(new Attendee(attendee.Name, attendee.Email));
            }
            foreach (SiebelAttendee attendee in appointment.OptionalAttendees)
            {
                a.OptionalAttendees.Add(new Attendee(attendee.Name, attendee.Email));
            }
            foreach (SiebelAttendee attendee in appointment.Resources)
            {
                a.Resources.Add(attendee.Name, attendee.Email);
            }
            
            if(a.Resources.Count >0)
            {
                a.SetExtendedProperty(MAPILocationDisplayName, (object)a.Resources[0].Name);
                if (appointment.MeetingLocation != null && appointment.MeetingLocation !="")
                    a.SetExtendedProperty(MAPILocationAnnotation, (object)appointment.MeetingLocation);
                else
                    a.RemoveExtendedProperty(MAPILocationAnnotation);
                a.SetExtendedProperty(MAPILocationUri, (object)a.Resources[0].Address);
                a.SetExtendedProperty(MAPILocationSource, (object)5);
            }
            else
            {
                a.Location = appointment.MeetingLocation;
            }

            a.Save(calendar.Id, SendInvitationsMode.SendOnlyToAll);
            //a.Update(ConflictResolutionMode.AlwaysOverwrite);
            EmailMessage message = new EmailMessage(exService); 
            message.Subject = "Создана встреча в календаре";
            MessageBody EmailMessageBody = new MessageBody (BodyType.HTML,"Уважаемый сотрудник,<br>В календаре создана встреча с клиентом: " + clientName + "<br>" +
                 changedSubject + "<br>" +
                "Ответственный: " + respName + "<br>" +
                "Место встречи: " + appointment.MeetingLocation + "<br>" +
                "Дата и время встречи: " + appointment.StartDate.ToString("dd.MM.yyyy HH:mm"));
            message.Body = EmailMessageBody;
            message.ToRecipients.Add(appointment.Owner.Email);
            message.Send();

            a.Load(new PropertySet(MAPICleanGlobalObjectId));
            a.TryGetProperty(MAPICleanGlobalObjectId, out oMAPICleanGlobalObjectId);
            return new SEIId(Convert.ToBase64String((Byte[])oMAPICleanGlobalObjectId), a.Id.UniqueId);
        }
        public static SEIId UpdateAppointment(SiebelAppointment appointment)
        {
            using (ThreadContext.Stacks["NDC"].Push("Updating appointments"))
            {

                SEIId res = new SEIId("", "");
                ExtendedPropertyDefinition MAPICleanGlobalObjectId = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.Meeting, 0x23, MapiPropertyType.Binary);
                ExtendedPropertyDefinition MAPILocationDisplayName = new ExtendedPropertyDefinition(new Guid("A719E259-2A9A-4FB8-BAB3-3A9F02970E4B"), "LocationDisplayName", MapiPropertyType.String);
                ExtendedPropertyDefinition MAPILocationUri = new ExtendedPropertyDefinition(new Guid("A719E259-2A9A-4FB8-BAB3-3A9F02970E4B"), "LocationUri", MapiPropertyType.String);
                ExtendedPropertyDefinition MAPILocationSource = new ExtendedPropertyDefinition(new Guid("A719E259-2A9A-4FB8-BAB3-3A9F02970E4B"), "LocationSource", MapiPropertyType.Integer);
                ExtendedPropertyDefinition MAPILocationAnnotation = new ExtendedPropertyDefinition(new Guid("A719E259-2A9A-4FB8-BAB3-3A9F02970E4B"), "LocationAnnotation", MapiPropertyType.String);
                object oMAPICleanGlobalObjectId;
                //object oMAPILocationDisplayName;
                //object oMAPILocationAnnotation;
                //object oMAPILocationSource;
                //object oMAPILocationUri;

                ExchangeService exService = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
                exService.Timeout = Properties.Settings.Default.ExchangeWSTimeout * 1000;
                exService.Url = new Uri(Properties.Settings.Default.ExchangeService);
                exService.Credentials = new NetworkCredential(Properties.Settings.Default.ExchangeLogin, Properties.Settings.Default.uExchangePassword, Properties.Settings.Default.ExchangeDomain);
                exService.TraceListener = new EWSTraceListener();
                exService.TraceFlags = TraceFlags.EwsRequest | TraceFlags.EwsResponse;
                exService.TraceEnabled = true;
                CalendarFolder calendar = CalendarFolder.Bind(exService, new FolderId(WellKnownFolderName.Calendar, appointment.Owner.Email));
                FindItemsResults<Item> foundItems = calendar.FindItems(new SearchFilter.IsEqualTo(MAPICleanGlobalObjectId, appointment.Id), new ItemView(1));
                if (foundItems.Items.Count > 0)
                {
                    /*foundItems.Items[0].Load(
                                    new PropertySet(
                                        BasePropertySet.FirstClassProperties,
                                        AppointmentSchema.Body,
                                        AppointmentSchema.AdjacentMeetingCount,
                                        AppointmentSchema.MimeContent,
                                        AppointmentSchema.UniqueBody,
                                        AppointmentSchema.Recurrence,
                                        CleanGlobalObjectId
                                    )
                                    { RequestedBodyType = BodyType.Text }
                                    );*/
                    Appointment apt = Appointment.Bind(exService,
                        foundItems.Items[0].Id,
                        new PropertySet(
                            BasePropertySet.FirstClassProperties,
                            AppointmentSchema.Body,
                            AppointmentSchema.AdjacentMeetingCount,
                            AppointmentSchema.MimeContent,
                            AppointmentSchema.UniqueBody,
                            AppointmentSchema.Recurrence
                            , MAPICleanGlobalObjectId
                            , MAPILocationDisplayName
                            , MAPILocationAnnotation
                            , MAPILocationSource
                            , MAPILocationUri
                            )
                        { RequestedBodyType = BodyType.Text }
                    );
                    string diffDebug = "Diffs:";
                    diffDebug += ((apt.Subject != appointment.Subject) ? "\n          " + apt.Subject + "->" + appointment.Subject : "");
                    //diffDebug += ((apt.Body != appointment.Body) ? "\n          " + apt.Body + "->" + appointment.Body : "");
                    diffDebug += ((apt.Start != appointment.StartDate) ? "\n          " + apt.Start.ToString() + "->" + appointment.StartDate.ToString() : "");
                    diffDebug += ((apt.End != appointment.EndDate) ? "\n          " + apt.End + "->" + appointment.EndDate.ToString() : "");
                    diffDebug += ((apt.ReminderMinutesBeforeStart != appointment.ReminderMinutesBeforeStart) ? "          " + apt.ReminderMinutesBeforeStart.ToString() + "->" + appointment.ReminderMinutesBeforeStart.ToString() : "");
                    diffDebug += ((apt.Location != appointment.MeetingLocation) ? "\n          " + apt.Location + "->" + appointment.MeetingLocation : "");

                    foreach (SiebelAttendee attendee in appointment.RequiredAttendees)
                        diffDebug += "\n          " + "+RA " + attendee.Name + "<" + attendee.Email + ">";
                    foreach (SiebelAttendee attendee in appointment.OptionalAttendees)
                        diffDebug += "\n          " + "+OA " + attendee.Name + "<" + attendee.Email + ">";
                    foreach (SiebelAttendee attendee in appointment.Resources)
                        diffDebug += "\n          " + "+RS " + attendee.Name + "<" + attendee.Email + ">";
                    foreach (Attendee attendee in apt.RequiredAttendees)
                        diffDebug += "\n          " + "-RA " + attendee.Name + "<" + attendee.Address + ">";
                    foreach (Attendee attendee in apt.OptionalAttendees)
                        diffDebug += "\n          " + "-OA " + attendee.Name + "<" + attendee.Address + ">";
                    foreach (Attendee attendee in apt.Resources)
                        diffDebug += "\n          " + "-RS " + attendee.Name + "<" + attendee.Address + ">";
                    //if (log.IsDebugEnabled) log.Info(diffDebug);
                    //Logger.Log(diffDebug);

                    int Offset = appointment.Subject.IndexOf("Ответственный:");
                    string clientName = String.Empty;
                    apt.Subject = appointment.Subject;
                    if (Offset != -1)
                    {
                        string exchDomain = Properties.Settings.Default.ExchangeDomain;
                        int OffsetClient = appointment.Subject.IndexOf("Клиент:");
                        if (OffsetClient != -1)
                        {
                            clientName = appointment.Subject.Substring(OffsetClient + 7, Offset - OffsetClient - 7);
                            string changedSubject = appointment.Subject.Remove(OffsetClient);
                            apt.Subject = changedSubject;
                        }
                        else
                        {
                            apt.Subject = appointment.Subject;
                        }

                        string respName = appointment.Subject.Substring(Offset + 14);
                        MessageBody mesBody = new MessageBody(BodyType.HTML, "Добрый день!<br>Назначена встреча с клиентом: " + clientName + "<br>Ответственный:" + respName +
                                             "<br>Место встречи: " + appointment.MeetingLocation + "<br>Дата и время встречи: " +
                                             appointment.StartDate.ToString("dd.MM.yyyy HH:mm") + "<br>Комментарий: Необходимо подтвердить дату и время встречи с клиентом за день<br>Ваш Сбербанк.");
                        apt.Body = mesBody;
                    }
                    else
                    {
                        apt.Subject = appointment.Subject;
                        //apt.Body = appointment.Body;
                    }

                    apt.Start = appointment.StartDate;
                    apt.End = appointment.EndDate;
                    apt.ReminderMinutesBeforeStart = appointment.ReminderMinutesBeforeStart;
                    foreach (SiebelAttendee attendee in appointment.RequiredAttendees)
                    {
                        //attendde is in exchange - nothing
                        //attendee isnt in exchange -insert
                        if (apt.RequiredAttendees.Where(f1 => f1.Address != null).Count(f => f.Address.Equals(attendee.Email, StringComparison.InvariantCultureIgnoreCase)) == 0)
                            apt.RequiredAttendees.Add(new Attendee(attendee.Name, attendee.Email));
                        //a.RequiredAttendees.Add(new Attendee(attendee.Name, attendee.Email));
                    }
                    List<string> atts4delete = new List<string>();
                    foreach (Attendee a in apt.RequiredAttendees.Where(f => f.Address != null))
                    {
                        //attendee isnt in siebel - delete;
                        if (appointment.RequiredAttendees.Count(f => f.Email.Equals(a.Address, StringComparison.InvariantCultureIgnoreCase)) == 0)
                            atts4delete.Add(a.Address);
                    }
                    foreach (string mail4delete in atts4delete)
                        apt.RequiredAttendees.Remove(apt.RequiredAttendees.Where(f1 => f1.Address != null).Single(f => f.Address.Equals(mail4delete, StringComparison.InvariantCultureIgnoreCase)));
                    while (apt.RequiredAttendees.Any(f => f.Address == null))
                    {
                        apt.RequiredAttendees.Remove(apt.RequiredAttendees.First(f => f.Address == null));
                    }

                    foreach (SiebelAttendee attendee in appointment.OptionalAttendees)
                    {
                        //attendde is in exchange - nothing
                        //attendee isnt in exchange -insert
                        if (apt.OptionalAttendees.Where(f1 => f1.Address != null).Count(f => f.Address.Equals(attendee.Email, StringComparison.InvariantCultureIgnoreCase)) == 0)
                            apt.OptionalAttendees.Add(new Attendee(attendee.Name, attendee.Email));
                        //a.RequiredAttendees.Add(new Attendee(attendee.Name, attendee.Email));
                    }
                    atts4delete.Clear();
                    foreach (Attendee a in apt.OptionalAttendees.Where(f => f.Address != null))
                    {
                        //attendee isnt in siebel - delete;
                        if (appointment.OptionalAttendees.Count(f => f.Email.Equals(a.Address, StringComparison.InvariantCultureIgnoreCase)) == 0)
                            atts4delete.Add(a.Address);
                    }
                    foreach (string mail4delete in atts4delete)
                        apt.OptionalAttendees.Remove(apt.OptionalAttendees.Where(f1 => f1.Address != null).Single(f => f.Address.Equals(mail4delete, StringComparison.InvariantCultureIgnoreCase)));
                    while (apt.OptionalAttendees.Any(f => f.Address == null))
                    {
                        apt.OptionalAttendees.Remove(apt.OptionalAttendees.First(f => f.Address == null));
                    }

                    foreach (SiebelAttendee attendee in appointment.Resources)
                    {
                        //attendde is in exchange - nothing
                        //attendee isnt in exchange -insert
                        if (apt.Resources.Where(f1 => f1.Address != null).Count(f => f.Address.Equals(attendee.Email, StringComparison.InvariantCultureIgnoreCase)) == 0)
                            apt.Resources.Add(new Attendee(attendee.Name, attendee.Email));
                        //a.RequiredAttendees.Add(new Attendee(attendee.Name, attendee.Email));
                    }
                    atts4delete.Clear();
                    foreach (Attendee a in apt.Resources.Where(f => f.Address != null))
                    {
                        //attendee isnt in siebel - delete;
                        if (appointment.Resources.Count(f => f.Email.Equals(a.Address, StringComparison.InvariantCultureIgnoreCase)) == 0)
                            atts4delete.Add(a.Address);
                    }
                    foreach (string mail4delete in atts4delete)
                        apt.Resources.Remove(apt.Resources.Where(f1 => f1.Address != null).Single(f => f.Address.Equals(mail4delete, StringComparison.InvariantCultureIgnoreCase)));
                    while (apt.Resources.Any(f => f.Address == null))
                    {
                        apt.Resources.Remove(apt.Resources.First(f => f.Address == null));
                    }

                    if (apt.Resources.Count > 0)
                    {
                        apt.SetExtendedProperty(MAPILocationDisplayName, (object)apt.Resources[0].Name);
                        if (appointment.MeetingLocation != null && appointment.MeetingLocation != "")
                            apt.SetExtendedProperty(MAPILocationAnnotation, (object)appointment.MeetingLocation);
                        else
                            apt.RemoveExtendedProperty(MAPILocationAnnotation);
                        apt.SetExtendedProperty(MAPILocationUri, (object)apt.Resources[0].Address);
                        apt.SetExtendedProperty(MAPILocationSource, (object)5);
                    }
                    else
                    {
                        apt.Location = appointment.MeetingLocation;
                    }

                    apt.Update(ConflictResolutionMode.AlwaysOverwrite, SendInvitationsOrCancellationsMode.SendOnlyToAll);

                    apt.TryGetProperty(MAPICleanGlobalObjectId, out oMAPICleanGlobalObjectId);

                    res.CGOID = Convert.ToBase64String((Byte[])oMAPICleanGlobalObjectId);
                    res.LocalId = apt.Id.UniqueId;
                }
                return res;
            }
        }

        public static bool CheckUserChanged(string UserEmail, string SyncToken)
        {
            ServicePointManager.ServerCertificateValidationCallback = delegate (object s, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) { return true; };
            using (ThreadContext.Stacks["NDC"].Push("CheckUserChanged"))
            {

                ExchangeService exService = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
                exService.Timeout = Properties.Settings.Default.ExchangeWSTimeout * 1000;
                exService.Url = new Uri(Properties.Settings.Default.ExchangeService);
                exService.Credentials = new NetworkCredential(Properties.Settings.Default.ExchangeLogin, Properties.Settings.Default.uExchangePassword, Properties.Settings.Default.ExchangeDomain);
                CalendarFolder calendar = new CalendarFolder(exService);
                try
                {
                    calendar = CalendarFolder.Bind(exService, new FolderId(WellKnownFolderName.Calendar, UserEmail));
                }
                catch (Exception e)
                {
                    log.Error(UserEmail + " - " + e.Message);
                    return false;
                }
                ChangeCollection<ItemChange> changes = exService.SyncFolderItems(calendar.Id, BasePropertySet.FirstClassProperties, null, 10, SyncFolderItemsScope.NormalItems, SyncToken);
                if (changes.Count > 0) return true;
                return false;
            }
        }
    }
}
