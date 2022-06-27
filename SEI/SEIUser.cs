using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SEI
{
    class SEIUser
    {
        private string direction;
        private string login;
        private string email;
        private DateTime syncStartDate;
        private string syncToken;
        private DateTime lastSyncDate;
        private string id;
        public string Direction { get { return direction; } }
        public string Login { get { return login; } }
        public string Email { get { return email; } }
        public DateTime SyncStartDate { get { return syncStartDate; } }
        public string SyncToken { get { return syncToken; } }
        public DateTime LastSyncDate { get { return lastSyncDate; } }
        public string Id { get { return id; } }

        public SEIUser(string login, string email, string direction, DateTime syncStartDate, string syncToken, DateTime? lastSyncDate, string id)
        {
            this.login = login;
            this.email = email;
            this.direction = direction;
            this.syncStartDate = syncStartDate;
            this.syncToken = syncToken;
            this.lastSyncDate = (lastSyncDate.HasValue?lastSyncDate.Value:syncStartDate);
            this.id = id;
        }
        public override string ToString()
        {
            return this.id + ";" + this.login + ";" + this.email + ";" + this.direction + ";" + this.syncStartDate.ToString() + ";" + this.syncToken + ";" + this.LastSyncDate.ToString();
        }
    }
}
