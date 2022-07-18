using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace RPA_API.Models
{
    public class OutlookResponse
    {
        public string responsecode { get; set; }
        public string responsemessage { get; set; }
        public messagedetails teamdetails { get; set; }
    }
    public class messagedetails
    {
        public string atmtechnical1 { get; set; }
        public string network { get; set; }
        public string sysadmin { get; set; }
        public string esupport { get; set; }
        public string basissupport1 { get; set; }
        public string basissupport2 { get; set; }
        public string atmtechnical2 { get; set; }
        public string datacentre { get; set; }
        public string consolidatedreport { get; set; }
    }
}
