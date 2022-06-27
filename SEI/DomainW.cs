using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.DirectoryServices;

namespace SEI
{
    class DomainW
    {
        public static List<String> GetRooms()
        {
            List<String> rooms = new List<String>();
            
            DirectoryEntry de = new DirectoryEntry(Properties.Settings.Default.LDAPAddress, Properties.Settings.Default.ExchangeLogin, Properties.Settings.Default.uExchangePassword);
            
            DirectorySearcher ds = new DirectorySearcher(de);
            ds.SearchScope = SearchScope.Subtree;
            ds.Filter = "(&(&(&(mail=*)(objectcategory=person)(objectclass=user)(msExchRecipientDisplayType=7))))";
            ds.PropertiesToLoad.Add("mail");
            SearchResultCollection src = ds.FindAll();
            foreach (SearchResult sr in src)
            {
                rooms.Add(sr.Properties["mail"][0].ToString());
            }

            return rooms;
        }
    }
}
