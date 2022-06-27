using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SEI
{
    /*class SiebelDeletedEmp
    {
        public string AttendeeEmail { get; set; }
        public string AppointmentId { get; set; }
        public string AppointmentSiebelId { get; set; }
        public string CommandId { get; set; }
        public SiebelDeletedEmp(string attendeeEmail, string appointmentId, string appointmentSiebelId, string commandId)
        {
            AttendeeEmail = attendeeEmail;
            AppointmentId = appointmentId;
            AppointmentSiebelId = appointmentSiebelId;
            CommandId = commandId;
        }
    }*/
    class SiebelAttendee
    {
        public string Email { get; set; }
        public string Id { get; set; }
        public string Name { get; set; }
        public string Response { get; set; }
        public SiebelAttendee (string id, string name, string email, string response)
        {
            Id = id;
            Name = name;
            Email = email;
            Response = response;
        }
        public SiebelAttendee()
        {

        }
    }
    
    class SiebelAppointment
    {
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public string Id { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }
        public string SiebelId { get; set; }
        public SiebelAttendee Owner { get; set; }
        public List<SiebelAttendee> RequiredAttendees { get; set; }
        public List<SiebelAttendee> OptionalAttendees { get; set; }
        public List<SiebelAttendee> Resources { get; set; }
        public DateTime LastUpdated { get; set; }
        public DateTime QueryDT { get; set; }
        public string OwnerId { get; set; }
        public string MeetingLocation { get; set; }
        public int ReminderMinutesBeforeStart { get; set; }

        public SiebelAppointment(string siebelId, string id, string subject, string body, DateTime startDate, DateTime endDate, DateTime lastUpdated, DateTime queryDT, string ownerId, string meetingLocation, int reminderMinutesBeforeStart)
        {
            StartDate = startDate;
            EndDate = endDate;
            SiebelId = siebelId;
            Id = id;
            Subject = subject;
            Body = body;
            Owner = new SiebelAttendee();
            RequiredAttendees = new List<SiebelAttendee>();
            OptionalAttendees = new List<SiebelAttendee>();
            Resources = new List<SiebelAttendee>();
            LastUpdated = lastUpdated;
            QueryDT = queryDT;
            OwnerId = ownerId;
            MeetingLocation = meetingLocation;
            ReminderMinutesBeforeStart = reminderMinutesBeforeStart;
        }

        public SiebelAppointment(string siebelId, string id, string ownerId)
        {
            SiebelId = siebelId;
            Id = id;
            OwnerId = ownerId;
        }
    }
}
