using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;
using log4net;

namespace SEI
{
    class SiebelResult
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public bool Status { get; set; }
        public string Id { get; set; }
        public SiebelResult(bool Status, string Id)
        {
            this.Status = Status;
            this.Id = Id;
        }
    }
    class SiebelWSW
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public static SiebelResult Delete(string localId, string siebelId)
        {
            using (ThreadContext.Stacks["NDC"].Push("Deleting appointment"))
            {

                SiebelResult result = new SiebelResult(false, "");
                try
                {
                    SiebelIntegrationSEI.SEI_spcIntegration_spcExchangeClient a1 = new SiebelIntegrationSEI.SEI_spcIntegration_spcExchangeClient(); 
                    a1.ClientCredentials.UserName.UserName = Properties.Settings.Default.SiebelLogin;
                    a1.ClientCredentials.UserName.Password = Properties.Settings.Default.uSiebelPassword;
                    a1.InnerChannel.OperationTimeout = new TimeSpan(0, 0, Properties.Settings.Default.SiebelWSDeleteTimeout);                  

                    SiebelIntegrationSEI.AppointmentTopElmt a2 = a1.DeleteAppointment(localId, siebelId);
                    if (a2.Appointment.ErrorCode == "0")
                    {
                        result.Status = true;
                    }
                    else
                    {
                        throw new ApplicationException("Siebel business exception: " + a2.Appointment.ErrorCode + "; " + a2.Appointment.ErrorText);
                    }

                    //Logger.Log("Delete Appointment End", Logger.LogSeverity.Debug);
                }
                catch (Exception e)
                {
                    log.Error(e.Message);
                    //Logger.Log(e.Message, Logger.LogSeverity.Error);
                    throw (e);
                }
                return result;
            }
        }
        public static SiebelResult InsertNew(Microsoft.Exchange.WebServices.Data.Appointment appointment, SEIUser user, List<String> rooms)
        {
            using (ThreadContext.Stacks["NDC"].Push("Inserting appointment"))
            {

                ExtendedPropertyDefinition CleanGlobalObjectId = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.Meeting, 0x23, MapiPropertyType.Binary);
                object CGOID;

                SiebelResult result = new SiebelResult(false, "");
                int i = 0;
                try
                {
                    //Logger.Log("Insert New Appointment Start", Logger.LogSeverity.Debug);

                    SiebelIntegrationSEI.SEI_spcIntegration_spcExchangeClient a1 = new SiebelIntegrationSEI.SEI_spcIntegration_spcExchangeClient();
                    a1.ClientCredentials.UserName.UserName = Properties.Settings.Default.SiebelLogin;
                    a1.ClientCredentials.UserName.Password = Properties.Settings.Default.uSiebelPassword;
                    a1.InnerChannel.OperationTimeout = new TimeSpan(0, 0, Properties.Settings.Default.SiebelWSTimeout);

                    SiebelIntegrationSEI.AppointmentTopElmt1 wsAppointment = new SiebelIntegrationSEI.AppointmentTopElmt1();
                    wsAppointment.Appointment = new SiebelIntegrationSEI.Appointment1();

                    wsAppointment.Appointment.AdjacentMeetingCount = appointment.AdjacentMeetingCount.ToString();
                    wsAppointment.Appointment.AdjacentMeetings = new SiebelIntegrationSEI.AdjacentMeetings();
                    //wsAppointment.Appointment.AllowNewTimeProposal = (appointment.AllowNewTimeProposal?"Y":"N");
                    wsAppointment.Appointment.AllowedResponseActions = appointment.AllowedResponseActions.ToString();
                    //wsAppointment.Appointment.AppointmentReplyTime = "";
                    wsAppointment.Appointment.AppointmentSequenceNumber = appointment.AppointmentSequenceNumber.ToString();
                    wsAppointment.Appointment.AppointmentState = appointment.AppointmentState.ToString();
                    wsAppointment.Appointment.AppointmentType = appointment.AppointmentType.ToString();
                    //wsAppointment.Appointment.ArchiveTag = new SiebelIntegrationSEI.ArchiveTag();
                    wsAppointment.Appointment.Attachments = new SiebelIntegrationSEI.Attachments();
                    wsAppointment.Appointment.Body = (appointment.Body.Text == null ? "" : (appointment.Body.Text.Length < 2000 ? appointment.Body.Text : (appointment.Body.Text.Substring(0, 1950) + "\r\n" + "...\r\n[More in Outlook]")));
                    wsAppointment.Appointment.Categories = new SiebelIntegrationSEI.Category[appointment.Categories.Count];
                    i = 0;
                    foreach (string category in appointment.Categories)
                    {
                        wsAppointment.Appointment.Categories[i] = new SiebelIntegrationSEI.Category();
                        wsAppointment.Appointment.Categories[i].Name = category;
                        wsAppointment.Appointment.Categories[i].Value = category;
                        i++;
                    }
                    //wsAppointment.Appointment.ConferenceType = appointment.ConferenceType.ToString();
                    //wsAppointment.Appointment.ConflictingMeetingCount = appointment.ConflictingMeetingCount.ToString();
                    wsAppointment.Appointment.ConflictingMeetings = new SiebelIntegrationSEI.ConflictingMeetings();
                    wsAppointment.Appointment.ConversationId = appointment.ConversationId.UniqueId;
                    wsAppointment.Appointment.Culture = appointment.Culture;
                    wsAppointment.Appointment.DateTimeCreated = appointment.DateTimeCreated.ToString("MM/dd/yyyy HH:mm:ss");
                    wsAppointment.Appointment.DateTimeReceived = appointment.DateTimeReceived.ToString("MM/dd/yyyy HH:mm:ss");
                    wsAppointment.Appointment.DateTimeSent = appointment.DateTimeSent.ToString("MM/dd/yyyy HH:mm:ss");

                    wsAppointment.Appointment.DeletedOccurrences = new SiebelIntegrationSEI.DeletedOccurrences();
                    wsAppointment.Appointment.DisplayCc = appointment.DisplayCc;
                    wsAppointment.Appointment.DisplayTo = appointment.DisplayTo;
                    wsAppointment.Appointment.Duration = new SiebelIntegrationSEI.Duration();
                    wsAppointment.Appointment.Duration.Days = appointment.Duration.Days.ToString();
                    wsAppointment.Appointment.Duration.Hours = appointment.Duration.Hours.ToString();
                    wsAppointment.Appointment.Duration.Minutes = appointment.Duration.Minutes.ToString();
                    wsAppointment.Appointment.Duration.Seconds = appointment.Duration.Seconds.ToString();
                    wsAppointment.Appointment.Duration.Milliseconds = appointment.Duration.Milliseconds.ToString();
                    wsAppointment.Appointment.EffectiveRights = appointment.EffectiveRights.ToString();
                    wsAppointment.Appointment.End = appointment.End.ToString("MM/dd/yyyy HH:mm:ss");
                    wsAppointment.Appointment.EndTimeZone = new SiebelIntegrationSEI.EndTimeZone();
                    wsAppointment.Appointment.EndTimeZone.Offset = new SiebelIntegrationSEI.Offset();
                    if (appointment.EndTimeZone != null)
                    {
                        wsAppointment.Appointment.EndTimeZone.Offset.Days = appointment.EndTimeZone.BaseUtcOffset.Days.ToString();
                        wsAppointment.Appointment.EndTimeZone.Offset.Hours = appointment.EndTimeZone.BaseUtcOffset.Hours.ToString();
                        wsAppointment.Appointment.EndTimeZone.Offset.Minutes = appointment.EndTimeZone.BaseUtcOffset.Minutes.ToString();
                        wsAppointment.Appointment.EndTimeZone.Offset.Seconds = appointment.EndTimeZone.BaseUtcOffset.Seconds.ToString();
                        wsAppointment.Appointment.EndTimeZone.Offset.Milliseconds = appointment.EndTimeZone.BaseUtcOffset.Milliseconds.ToString();
                    }
                    //wsAppointment.Appointment.EnhancedLocation = new SiebelIntegrationSEI.EnhancedLocation();
                    //wsAppointment.Appointment.EntityExtractionResult = new SiebelIntegrationSEI.EntityExtractionResult();
                    wsAppointment.Appointment.ExtendedProperties = new SiebelIntegrationSEI.ExtendedProperty[0];
                    wsAppointment.Appointment.FirstOccurrence = new SiebelIntegrationSEI.FirstOccurrence();
                    //wsAppointment.Appointment.Flag = new SiebelIntegrationSEI.Flag();
                    wsAppointment.Appointment.HasAttachments = (appointment.HasAttachments ? "Y" : "N");
                    if (appointment.ICalDateTimeStamp.HasValue) wsAppointment.Appointment.ICalDateTimeStamp = appointment.ICalDateTimeStamp.Value.ToString("MM/dd/yyyy HH:mm:ss");
                    if (appointment.ICalRecurrenceId.HasValue) wsAppointment.Appointment.ICalRecurrenceId = appointment.ICalRecurrenceId.Value.ToString("MM/dd/yyyy HH:mm:ss");
                    wsAppointment.Appointment.ICalUid = appointment.ICalUid;
                    //wsAppointment.Appointment.IconIndex = new SiebelIntegrationSEI.IconIndex();

                    wsAppointment.Appointment.LocalId = appointment.Id.UniqueId;

                    appointment.TryGetProperty(CleanGlobalObjectId, out CGOID);
                    wsAppointment.Appointment.Id = Convert.ToBase64String((Byte[])CGOID);

                    wsAppointment.Appointment.Importance = appointment.Importance.ToString();
                    wsAppointment.Appointment.InReplyTo = appointment.InReplyTo;
                    //wsAppointment.Appointment.InstanceKey = System.Text.UTF8Encoding.UTF8.GetString(appointment.InstanceKey);
                    wsAppointment.Appointment.InternetMessageHeaders = new SiebelIntegrationSEI.InternetMessageHeaders();
                    wsAppointment.Appointment.IsAllDayEvent = (appointment.IsAllDayEvent ? "Y" : "N");
                    wsAppointment.Appointment.IsAssociated = (appointment.IsAssociated ? "Y" : "N");
                    wsAppointment.Appointment.IsAttachment = (appointment.IsAttachment ? "Y" : "N");
                    wsAppointment.Appointment.IsCancelled = (appointment.IsCancelled ? "Y" : "N");
                    wsAppointment.Appointment.IsDirty = (appointment.IsDirty ? "Y" : "N");
                    wsAppointment.Appointment.IsDraft = (appointment.IsDraft ? "Y" : "N");
                    wsAppointment.Appointment.IsFromMe = (appointment.IsFromMe ? "Y" : "N");
                    wsAppointment.Appointment.IsMeeting = (appointment.IsMeeting ? "Y" : "N");
                    wsAppointment.Appointment.IsNew = (appointment.IsNew ? "Y" : "N");
                    //wsAppointment.Appointment.IsOnlineMeeting = (appointment.IsOnlineMeeting ? "Y" : "N");
                    wsAppointment.Appointment.IsRecurring = (appointment.IsRecurring ? "Y" : "N");
                    wsAppointment.Appointment.IsReminderSet = (appointment.IsReminderSet ? "Y" : "N");
                    wsAppointment.Appointment.IsResend = (appointment.IsResend ? "Y" : "N");
                    wsAppointment.Appointment.IsResponseRequested = (appointment.IsResponseRequested ? "Y" : "N");
                    wsAppointment.Appointment.IsSubmitted = (appointment.IsSubmitted ? "Y" : "N");
                    wsAppointment.Appointment.IsUnmodified = (appointment.IsUnmodified ? "Y" : "N");
                    wsAppointment.Appointment.ItemClass = appointment.ItemClass;
                    //wsAppointment.Appointment.JoinOnlineMeetingUrl = appointment.JoinOnlineMeetingUrl;
                    wsAppointment.Appointment.LastModifiedName = appointment.LastModifiedName;
                    wsAppointment.Appointment.LastModifiedTime = appointment.LastModifiedTime.ToString("MM/dd/yyyy HH:mm:ss");
                    wsAppointment.Appointment.LastOccurrence = new SiebelIntegrationSEI.LastOccurrence();
                    wsAppointment.Appointment.LegacyFreeBusyStatus = appointment.LegacyFreeBusyStatus.ToString();
                    wsAppointment.Appointment.Location = appointment.Location;
                    wsAppointment.Appointment.MeetingRequestWasSent = (appointment.MeetingRequestWasSent ? "Y" : "N");
                    wsAppointment.Appointment.MeetingWorkspaceUrl = appointment.MeetingWorkspaceUrl;
                    wsAppointment.Appointment.MimeContent = new SiebelIntegrationSEI.MimeContent();
                    wsAppointment.Appointment.MimeContent.StringElement = appointment.MimeContent.ToString();
                    wsAppointment.Appointment.ModifiedOccurrences = new SiebelIntegrationSEI.ModifiedOccurrences();
                    wsAppointment.Appointment.MyResponseType = appointment.MyResponseType.ToString();
                    wsAppointment.Appointment.NetShowUrl = appointment.NetShowUrl;
                    //wsAppointment.Appointment.NormalizedBody = appointment.NormalizedBody.Text;
                    //wsAppointment.Appointment.OnlineMeetingSettings = new SiebelIntegrationSEI.OnlineMeetingSettings();
                    wsAppointment.Appointment.OptionalAttendees = new SiebelIntegrationSEI.Attendee[appointment.OptionalAttendees.Where(f => !rooms.Exists(f1 => f1.Equals(f.Address, StringComparison.InvariantCultureIgnoreCase))).Count()];
                    i = 0;
                    foreach (Microsoft.Exchange.WebServices.Data.Attendee attendee in appointment.OptionalAttendees.Where(f => !rooms.Exists(f1 => f1.Equals(f.Address, StringComparison.InvariantCultureIgnoreCase))))
                    {
                        wsAppointment.Appointment.OptionalAttendees[i] = new SiebelIntegrationSEI.Attendee();
                        wsAppointment.Appointment.OptionalAttendees[i].Address = attendee.Address;
                        if (attendee.Id != null) wsAppointment.Appointment.OptionalAttendees[i].Id = attendee.Id.UniqueId;
                        wsAppointment.Appointment.OptionalAttendees[i].MailboxType = (attendee.MailboxType.HasValue ? attendee.MailboxType.Value.ToString() : "");
                        wsAppointment.Appointment.OptionalAttendees[i].Name = attendee.Name;
                        wsAppointment.Appointment.OptionalAttendees[i].LastResponseTime = (attendee.LastResponseTime.HasValue ? attendee.LastResponseTime.Value.ToString("MM/dd/yyyy HH:mm:ss") : "");
                        wsAppointment.Appointment.OptionalAttendees[i].ResponseType = attendee.ResponseType.ToString();
                        wsAppointment.Appointment.OptionalAttendees[i].RoutingType = attendee.RoutingType;
                        i++;
                    }
                    wsAppointment.Appointment.Organizer = new SiebelIntegrationSEI.Organizer();
                    wsAppointment.Appointment.Organizer.Address = appointment.Organizer.Address;
                    if (appointment.Organizer.Id != null) wsAppointment.Appointment.Organizer.Id = appointment.Organizer.Id.UniqueId;
                    wsAppointment.Appointment.Organizer.MailboxType = (appointment.Organizer.MailboxType.HasValue ? appointment.Organizer.MailboxType.Value.ToString() : "");
                    wsAppointment.Appointment.Organizer.Name = appointment.Organizer.Name;
                    wsAppointment.Appointment.Organizer.RoutingType = appointment.Organizer.RoutingType;
                    //wsAppointment.Appointment.OriginalStart = appointment.OriginalStart.ToString("MM/dd/yyyy HH:mm:ss");
                    wsAppointment.Appointment.ParentFolderId = appointment.ParentFolderId.UniqueId;
                    //wsAppointment.Appointment.PolicyTag = new SiebelIntegrationSEI.PolicyTag();
                    //wsAppointment.Appointment.Preview = new SiebelIntegrationSEI.Preview();
                    wsAppointment.Appointment.Recurrence = new SiebelIntegrationSEI.Recurrence();
                    wsAppointment.Appointment.ReminderDueBy = appointment.ReminderDueBy.ToString("MM/dd/yyyy HH:mm:ss");
                    wsAppointment.Appointment.ReminderMinutesBeforeStart = appointment.ReminderMinutesBeforeStart.ToString();
                    wsAppointment.Appointment.RequiredAttendees = new SiebelIntegrationSEI.Attendee[appointment.RequiredAttendees.Where(f => !rooms.Exists(f1 => f1.Equals(f.Address, StringComparison.InvariantCultureIgnoreCase))).Count()];
                    i = 0;
                    foreach (Microsoft.Exchange.WebServices.Data.Attendee attendee in appointment.RequiredAttendees.Where(f => !rooms.Exists(f1 => f1.Equals(f.Address, StringComparison.InvariantCultureIgnoreCase))))
                    {
                        wsAppointment.Appointment.RequiredAttendees[i] = new SiebelIntegrationSEI.Attendee();
                        wsAppointment.Appointment.RequiredAttendees[i].Address = attendee.Address;
                        if (attendee.Id != null) wsAppointment.Appointment.RequiredAttendees[i].Id = attendee.Id.UniqueId;
                        wsAppointment.Appointment.RequiredAttendees[i].MailboxType = (attendee.MailboxType.HasValue ? attendee.MailboxType.Value.ToString() : "");
                        wsAppointment.Appointment.RequiredAttendees[i].Name = attendee.Name;
                        wsAppointment.Appointment.RequiredAttendees[i].LastResponseTime = (attendee.LastResponseTime.HasValue ? attendee.LastResponseTime.Value.ToString("MM/dd/yyyy HH:mm:ss") : "");
                        wsAppointment.Appointment.RequiredAttendees[i].ResponseType = attendee.ResponseType.ToString();
                        wsAppointment.Appointment.RequiredAttendees[i].RoutingType = attendee.RoutingType;
                        i++;
                    }
                    wsAppointment.Appointment.Resources = new SiebelIntegrationSEI.Attendee[appointment.Resources.Count + appointment.OptionalAttendees.Count(f => rooms.Exists(f1 => f1.Equals(f.Address, StringComparison.InvariantCultureIgnoreCase))) + appointment.RequiredAttendees.Count(f => rooms.Exists(f1 => f1.Equals(f.Address, StringComparison.InvariantCultureIgnoreCase)))];
                    i = 0;
                    foreach (Microsoft.Exchange.WebServices.Data.Attendee attendee in appointment.Resources)
                    {
                        wsAppointment.Appointment.Resources[i] = new SiebelIntegrationSEI.Attendee();
                        wsAppointment.Appointment.Resources[i].Address = attendee.Address;
                        if (attendee.Id != null) wsAppointment.Appointment.Resources[i].Id = attendee.Id.UniqueId;
                        wsAppointment.Appointment.Resources[i].MailboxType = (attendee.MailboxType.HasValue ? attendee.MailboxType.Value.ToString() : "");
                        wsAppointment.Appointment.Resources[i].Name = attendee.Name;
                        wsAppointment.Appointment.Resources[i].LastResponseTime = (attendee.LastResponseTime.HasValue ? attendee.LastResponseTime.Value.ToString("MM/dd/yyyy HH:mm:ss") : "");
                        wsAppointment.Appointment.Resources[i].ResponseType = attendee.ResponseType.ToString();
                        wsAppointment.Appointment.Resources[i].RoutingType = attendee.RoutingType;
                        i++;
                    }
                    foreach (Microsoft.Exchange.WebServices.Data.Attendee attendee in appointment.RequiredAttendees.Where(f => rooms.Exists(f1 => f1.Equals(f.Address, StringComparison.InvariantCultureIgnoreCase))))
                    {
                        wsAppointment.Appointment.Resources[i] = new SiebelIntegrationSEI.Attendee();
                        wsAppointment.Appointment.Resources[i].Address = attendee.Address;
                        if (attendee.Id != null) wsAppointment.Appointment.Resources[i].Id = attendee.Id.UniqueId;
                        wsAppointment.Appointment.Resources[i].MailboxType = (attendee.MailboxType.HasValue ? attendee.MailboxType.Value.ToString() : "");
                        wsAppointment.Appointment.Resources[i].Name = attendee.Name;
                        wsAppointment.Appointment.Resources[i].LastResponseTime = (attendee.LastResponseTime.HasValue ? attendee.LastResponseTime.Value.ToString("MM/dd/yyyy HH:mm:ss") : "");
                        wsAppointment.Appointment.Resources[i].ResponseType = attendee.ResponseType.ToString();
                        wsAppointment.Appointment.Resources[i].RoutingType = attendee.RoutingType;
                        i++;
                    }
                    foreach (Microsoft.Exchange.WebServices.Data.Attendee attendee in appointment.OptionalAttendees.Where(f => rooms.Exists(f1 => f1.Equals(f.Address, StringComparison.InvariantCultureIgnoreCase))))
                    {
                        wsAppointment.Appointment.Resources[i] = new SiebelIntegrationSEI.Attendee();
                        wsAppointment.Appointment.Resources[i].Address = attendee.Address;
                        if (attendee.Id != null) wsAppointment.Appointment.Resources[i].Id = attendee.Id.UniqueId;
                        wsAppointment.Appointment.Resources[i].MailboxType = (attendee.MailboxType.HasValue ? attendee.MailboxType.Value.ToString() : "");
                        wsAppointment.Appointment.Resources[i].Name = attendee.Name;
                        wsAppointment.Appointment.Resources[i].LastResponseTime = (attendee.LastResponseTime.HasValue ? attendee.LastResponseTime.Value.ToString("MM/dd/yyyy HH:mm:ss") : "");
                        wsAppointment.Appointment.Resources[i].ResponseType = attendee.ResponseType.ToString();
                        wsAppointment.Appointment.Resources[i].RoutingType = attendee.RoutingType;
                        i++;
                    }
                    //if (appointment.RetentionDate.HasValue) wsAppointment.Appointment.RetentionDate = appointment.RetentionDate.Value.ToString("MM/dd/yyyy HH:mm:ss");
                    wsAppointment.Appointment.Schema = appointment.Schema.ToString();
                    wsAppointment.Appointment.Sensitivity = appointment.Sensitivity.ToString();
                    wsAppointment.Appointment.Size = appointment.Size.ToString();
                    wsAppointment.Appointment.Start = appointment.Start.ToString("MM/dd/yyyy HH:mm:ss");
                    wsAppointment.Appointment.StartTimeZone = new SiebelIntegrationSEI.StartTimeZone();
                    wsAppointment.Appointment.StartTimeZone.Offset = new SiebelIntegrationSEI.Offset();
                    if (appointment.StartTimeZone != null)
                    {
                        wsAppointment.Appointment.StartTimeZone.Offset.Days = appointment.StartTimeZone.BaseUtcOffset.Days.ToString();
                        wsAppointment.Appointment.StartTimeZone.Offset.Hours = appointment.StartTimeZone.BaseUtcOffset.Hours.ToString();
                        wsAppointment.Appointment.StartTimeZone.Offset.Minutes = appointment.StartTimeZone.BaseUtcOffset.Minutes.ToString();
                        wsAppointment.Appointment.StartTimeZone.Offset.Seconds = appointment.StartTimeZone.BaseUtcOffset.Seconds.ToString();
                        wsAppointment.Appointment.StartTimeZone.Offset.Milliseconds = appointment.StartTimeZone.BaseUtcOffset.Milliseconds.ToString();
                    }
                    if (appointment.StoreEntryId != null) wsAppointment.Appointment.StoreEntryId = System.Text.UTF8Encoding.UTF8.GetString(appointment.StoreEntryId);
                    wsAppointment.Appointment.Subject = appointment.Subject;
                    //wsAppointment.Appointment.TextBody = appointment.TextBody.Text;
                    wsAppointment.Appointment.TimeZone = appointment.TimeZone;
                    wsAppointment.Appointment.UniqueBody = new SiebelIntegrationSEI.UniqueBody();
                    if (appointment.UniqueBody != null)
                    {
                        wsAppointment.Appointment.UniqueBody.BodyType = appointment.UniqueBody.BodyType.ToString();
                        wsAppointment.Appointment.UniqueBody.IsTruncated = (appointment.UniqueBody.IsTruncated ? "Y" : "N");
                        wsAppointment.Appointment.UniqueBody.Text = appointment.UniqueBody.Text;
                    }
                    wsAppointment.Appointment.WebClientEditFormQueryString = appointment.WebClientEditFormQueryString;
                    wsAppointment.Appointment.WebClientReadFormQueryString = appointment.WebClientReadFormQueryString;
                    wsAppointment.Appointment.When = appointment.When;

                    wsAppointment.Appointment.IsOrganized = (user.Email.Equals(appointment.Organizer.Address, StringComparison.InvariantCultureIgnoreCase) ? "Y" : "N");
                    wsAppointment.Appointment.RequestorId = user.Id;
                    wsAppointment.Appointment.RequestorLogin = user.Login;

                    SiebelIntegrationSEI.AppointmentTopElmt a2 = a1.NewAppointment(wsAppointment);

                    if (a2.Appointment.ErrorCode == "0")
                    {
                        result.Status = true;
                        result.Id = a2.Appointment.SiebelId;
                    }
                    else
                    {
                        log.Error("Siebel business exception: " + a2.Appointment.ErrorCode + "; " + a2.Appointment.ErrorText);
                        throw new ApplicationException("Siebel business exception: " + a2.Appointment.ErrorCode + "; " + a2.Appointment.ErrorText);
                    }

                    //Logger.Log("Insert New Appointment End", Logger.LogSeverity.Debug);
                }
                catch (Exception e)
                {
                    log.Error(e.Message);
                    //Logger.Log(e.Message, Logger.LogSeverity.Error);
                    throw (e);
                }
                return result;
            }
        }
        public static void DeleteSEIUser(string userId)
        {
            using (ThreadContext.Stacks["NDC"].Push("Deleting SEI user"))
            {
                try
                {
                    SiebelIntegrationSEI.SEI_spcIntegration_spcExchangeClient client = new SiebelIntegrationSEI.SEI_spcIntegration_spcExchangeClient();
                    client.ClientCredentials.UserName.UserName = Properties.Settings.Default.SiebelLogin;
                    client.ClientCredentials.UserName.Password = Properties.Settings.Default.uSiebelPassword;
                    client.InnerChannel.OperationTimeout = new TimeSpan(0, 0, Properties.Settings.Default.SiebelWSDeleteTimeout);
                    SEI.SiebelIntegrationSEI.SEIUserOutputTopElmt result = client.DeleteSEIUser(userId);
                    if (result.SEIUserOutput.ErrorCode == "0")
                    {
                        if (log.IsInfoEnabled) log.Info("Deleted SEI user - " + userId);
                    }
                    else
                    {
                        if (log.IsInfoEnabled) log.Info(userId + " not deleted: " + result.SEIUserOutput.ErrorText);
                    }
                    
                }
                catch (Exception e)
                {
                    log.Error(e.Message);
                }
            }
        }
        public static void ChangeActStatus(string ActionId, string State)
        {
            using (ThreadContext.Stacks["NDC"].Push("Setting parent activity state"))
            {
                try
                {
                    SiebelIntegrationSEI.SEI_spcIntegration_spcExchangeClient client = new SiebelIntegrationSEI.SEI_spcIntegration_spcExchangeClient();
                    client.ClientCredentials.UserName.UserName = Properties.Settings.Default.SiebelLogin;
                    client.ClientCredentials.UserName.Password = Properties.Settings.Default.uSiebelPassword;
                    client.InnerChannel.OperationTimeout = new TimeSpan(0, 0, Properties.Settings.Default.SiebelWSDeleteTimeout);
                    SEI.SiebelIntegrationSEI.SEIUserOutputTopElmt result = client.ChangeActStatus(ActionId, State);
                    if (result.SEIUserOutput.ErrorCode == "0")
                    {
                        if (log.IsInfoEnabled) log.Info("Parent activity state was set to " + State + " for appointment " + ActionId);
                    }
                    else
                    {
                        if (log.IsInfoEnabled) log.Info("Parent activity state for " + ActionId + " was not set: " + result.SEIUserOutput.ErrorText);
                    }

                }
                catch (Exception e)
                {
                    log.Error(e.Message);
                }

            }
        }
        public static void InsertSEIUser(string userId, string Email)
        {
            using (ThreadContext.Stacks["NDC"].Push("Inserting SEI user"))
            {

                string result = string.Empty;
                try
                {
                    SiebelIntegrationSEI.SEI_spcIntegration_spcExchangeClient client = new SiebelIntegrationSEI.SEI_spcIntegration_spcExchangeClient();
                    client.ClientCredentials.UserName.UserName = Properties.Settings.Default.SiebelLogin;
                    client.ClientCredentials.UserName.Password = Properties.Settings.Default.uSiebelPassword;
                    client.InnerChannel.OperationTimeout = new TimeSpan(0, 0, Properties.Settings.Default.SiebelWSDeleteTimeout);
                    SEI.SiebelIntegrationSEI.SEIUserOutputTopElmt result1 = client.InsertSEIUser(userId, Email);
                    if (result1.SEIUserOutput.ErrorCode == "0")
                    {
                        if (log.IsInfoEnabled) log.Info("Inserted SEI user - " + userId);
                    }
                    else
                    {
                        if (log.IsInfoEnabled) log.Info(userId + " not inserted: " + result1.SEIUserOutput.ErrorText);
                    }
                }
                catch (Exception e)
                {
                    log.Error(e.Message);
                }
            }
        }

        public static bool CheckExchangeEmail(string Email)
        {

            return false;
        }

    }


}
