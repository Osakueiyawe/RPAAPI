using Microsoft.Extensions.Configuration;
using RPA_API.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace RPA_API.Methods
{
    public class ConnectToOutlook : IConnectToOutlook
    {
        public IConfiguration Configuration { get; set; }
        public ConnectToOutlook(IConfiguration configuration)
        {
            Configuration = configuration;
        }
        public async Task<OutlookResponse> Outlookdetails(OutlookRequest userdetails)
        {
            var outlookresponse = new OutlookResponse();
            var messagedetails = new messagedetails();
            try
            {
                Outlook.Application oApp = new Outlook.Application();
                Outlook.NameSpace oNS = oApp.GetNamespace("mapi");
                oNS.Logon(userdetails.username, userdetails.password, false, true);
                Outlook.MAPIFolder oInbox = oNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                Outlook.Items oItems = oInbox.Items;
                var date = DateTime.Now;
                var expectedemails = Convert.ToInt32(Configuration.GetSection("noofexpectedemails").Value);
                oItems = oItems.Restrict("[ReceivedTime] > '" + date.ToString("dd/MM/yyyy") + "'");
                oItems = oItems.Restrict("@SQL=\"urn:schemas:httpmail:subject\" like '%" + Configuration.GetSection("emailsubjectsearch").Value + "%'");
                int count = oItems.Count;
                LogError.Errhandler($"Successfully retrieved {count} emails that fulfill the required conditions");
                List<Outlook.MailItem> mails = new List<Outlook.MailItem>();
                foreach (Outlook.MailItem item in oItems)
                {
                    if (item is Outlook.MailItem)
                    {
                        Outlook.MailItem mail = (Outlook.MailItem)item;
                        mails.Add(mail);
                        var subject = mail.Subject;
                        subject = subject.ToUpper();
                        var searchkey = Configuration.GetSection("emailsubjectsearch").Value;
                        if (subject.Contains(searchkey))
                        {
                            count = count++;
                        }
                    }
                }
                List<String> teams = new List<string>();
                teams.Add(Configuration.GetSection("TeamEmails").GetSection("atmtechnical1").Value);
                teams.Add(Configuration.GetSection("TeamEmails").GetSection("network").Value);
                teams.Add(Configuration.GetSection("TeamEmails").GetSection("sysadmin").Value);
                teams.Add(Configuration.GetSection("TeamEmails").GetSection("esupport").Value);
                teams.Add(Configuration.GetSection("TeamEmails").GetSection("basissupport1").Value);
                teams.Add(Configuration.GetSection("TeamEmails").GetSection("basissupport2").Value);
                teams.Add(Configuration.GetSection("TeamEmails").GetSection("atmtechnical2").Value);
                teams.Add(Configuration.GetSection("TeamEmails").GetSection("datacentre").Value);
                teams.Add(Configuration.GetSection("TeamEmails").GetSection("consolidatedreport").Value);

                
                if (Directory.Exists(Configuration.GetSection("filelocation").Value))
                {
                    Directory.Delete(Configuration.GetSection("filelocation").Value, true);
                }

                LogError.Errhandler("About to upload attachments of each received emails...");
                foreach (Outlook.MailItem email in mails)
                {
                    
                    if (email.Subject.ToUpper().Contains(Configuration.GetSection("TeamEmails").GetSection("atmtechnical1").Value) && email.Attachments.Count == 1)
                    {
                        teams.Remove(Configuration.GetSection("TeamEmails").GetSection("atmtechnical1").Value);
                        var atmtechnical1 = Configuration.GetSection("filelocation").Value + "/ATMTECHNICAL1/";
                        var fullatmtechnical1 = atmtechnical1 + email.Attachments[1].FileName;
                        messagedetails.atmtechnical1 = fullatmtechnical1;
                        if (!Directory.Exists(atmtechnical1))
                        {
                            Directory.CreateDirectory(atmtechnical1);
                        }
                        email.Attachments[1].SaveAsFile(fullatmtechnical1);
                        LogError.Errhandler("Saved Attachment for ATM Technical 1");
                    }
                    else if (email.Subject.ToUpper().Contains(Configuration.GetSection("TeamEmails").GetSection("network").Value) && email.Attachments.Count == 1)
                    {
                        teams.Remove(Configuration.GetSection("TeamEmails").GetSection("network").Value);
                        var network = Configuration.GetSection("filelocation").Value + "/NETWORK/";
                        var fullnetwork = Configuration.GetSection("filelocation").Value + "/NETWORK/" + email.Attachments[1].FileName;
                        messagedetails.network = fullnetwork;
                        if (!Directory.Exists(network))
                        {
                            Directory.CreateDirectory(network);
                        }
                        email.Attachments[1].SaveAsFile(fullnetwork);
                        LogError.Errhandler("Saved Attachment for the network Team");
                    }
                    else if (email.Subject.ToUpper().Contains(Configuration.GetSection("TeamEmails").GetSection("sysadmin").Value) && email.Attachments.Count == 1)
                    {
                        teams.Remove(Configuration.GetSection("TeamEmails").GetSection("sysadmin").Value);
                        var sysadmin = Configuration.GetSection("filelocation").Value + "/SYSADMIN/";
                        var fullsysadmin = Configuration.GetSection("filelocation").Value + "/SYSADMIN/" + email.Attachments[1].FileName;
                        messagedetails.sysadmin = fullsysadmin;
                        if (!Directory.Exists(sysadmin))
                        {
                            Directory.CreateDirectory(sysadmin);
                        }
                        email.Attachments[1].SaveAsFile(fullsysadmin);
                        LogError.Errhandler("Saved attachment for sysadmin");
                    }
                    else if (email.Subject.ToUpper().Contains(Configuration.GetSection("TeamEmails").GetSection("esupport").Value) && email.Attachments.Count == 1)
                    {
                        teams.Remove(Configuration.GetSection("TeamEmails").GetSection("esupport").Value);
                        var esupport = Configuration.GetSection("filelocation").Value + "/ESUPPORT/";
                        var fullesupport = Configuration.GetSection("filelocation").Value + "/ESUPPORT/" + email.Attachments[1].FileName;
                        messagedetails.esupport = fullesupport;
                        if (!Directory.Exists(esupport))
                        {
                            Directory.CreateDirectory(esupport);
                        }
                        email.Attachments[1].SaveAsFile(fullesupport);
                        LogError.Errhandler("Saved attachment for esupport");
                    }
                    else if (email.Subject.ToUpper().Contains(Configuration.GetSection("TeamEmails").GetSection("basissupport1").Value) && email.Attachments.Count == 1)
                    {
                        teams.Remove(Configuration.GetSection("TeamEmails").GetSection("basissupport1").Value);
                        var basissupport1 = Configuration.GetSection("filelocation").Value + "/BASISSUPPORT1/";
                        var fullbasissupport1 = Configuration.GetSection("filelocation").Value + "/BASISSUPPORT1/" + email.Attachments[1].FileName;
                        messagedetails.basissupport1 = fullbasissupport1;
                        if (!Directory.Exists(basissupport1))
                        {
                            Directory.CreateDirectory(basissupport1);
                        }
                        email.Attachments[1].SaveAsFile(fullbasissupport1);
                        LogError.Errhandler("Saved attachment for basis support 1");
                    }
                    else if (email.Subject.ToUpper().Contains(Configuration.GetSection("TeamEmails").GetSection("basissupport2").Value) && email.Attachments.Count == 1)
                    {
                        teams.Remove(Configuration.GetSection("TeamEmails").GetSection("basissupport2").Value);
                        var basissupport2 = Configuration.GetSection("filelocation").Value + "/BASISSUPPORT2/" + email.Attachments[1].FileName;
                        var fullbasissupport2 = Configuration.GetSection("filelocation").Value + "/BASISSUPPORT2/" + email.Attachments[1].FileName;
                        messagedetails.basissupport2 = fullbasissupport2;
                        if (!Directory.Exists(basissupport2))
                        {
                            Directory.CreateDirectory(basissupport2);
                        }
                        email.Attachments[1].SaveAsFile(fullbasissupport2);
                        LogError.Errhandler("Saved attachment for basis support 2");
                    }
                    else if (email.Subject.ToUpper().Contains(Configuration.GetSection("TeamEmails").GetSection("atmtechnical2").Value) && email.Attachments.Count == 1)
                    {
                        teams.Remove(Configuration.GetSection("TeamEmails").GetSection("atmtechnical2").Value);
                        var atmtechnical2 = Configuration.GetSection("filelocation").Value + "/ATMTECHNICAL2/";
                        var fullatmtechnical2 = Configuration.GetSection("filelocation").Value + "/ATMTECHNICAL2/" + email.Attachments[1].FileName;
                        messagedetails.atmtechnical2 = fullatmtechnical2;
                        if (!Directory.Exists(atmtechnical2))
                        {
                            Directory.CreateDirectory(atmtechnical2);
                        }
                        email.Attachments[1].SaveAsFile(fullatmtechnical2);
                        LogError.Errhandler("Saved attachment for ATM Technical 2");
                    }
                    else if (email.Subject.ToUpper().Contains(Configuration.GetSection("TeamEmails").GetSection("datacentre").Value) && email.Attachments.Count == 1)
                    {
                        teams.Remove(Configuration.GetSection("TeamEmails").GetSection("datacentre").Value);
                        var datacentre = Configuration.GetSection("filelocation").Value + "/DATACENTRE/";
                        var fulldatacentre = Configuration.GetSection("filelocation").Value + "/DATACENTRE/" + email.Attachments[1].FileName;
                        messagedetails.datacentre = fulldatacentre;
                        if (!Directory.Exists(datacentre))
                        {
                            Directory.CreateDirectory(datacentre);
                        }
                        email.Attachments[1].SaveAsFile(fulldatacentre);
                        LogError.Errhandler("Saved attachment for data centre");
                    }
                    else if (email.Subject.ToUpper().Contains(Configuration.GetSection("TeamEmails").GetSection("consolidatedreport").Value) && email.Attachments.Count == 1)
                    {
                        teams.Remove(Configuration.GetSection("TeamEmails").GetSection("consolidatedreport").Value);
                        var consolidatedreport = Configuration.GetSection("filelocation").Value + "/CONSOLIDATEDREPORT/";
                        var fullconsolidatedreport = Configuration.GetSection("filelocation").Value + "/CONSOLIDATEDREPORT/" + email.Attachments[1].FileName;
                        messagedetails.consolidatedreport = fullconsolidatedreport;
                        if (!Directory.Exists(consolidatedreport))
                        {
                            Directory.CreateDirectory(consolidatedreport);
                        }
                        email.Attachments[1].SaveAsFile(fullconsolidatedreport);
                        LogError.Errhandler("Saved attachment for consolidated Report");
                    }                   

                }
                outlookresponse.teamdetails = messagedetails;
                if (teams.Count > 0)
                {
                    //outlookresponse = null;
                    outlookresponse.responsecode = "00";
                    string listofteams = string.Join(",", teams);
                    outlookresponse.responsemessage = "The following teams have not sent their emails: " + listofteams;
                }
                if (count >= expectedemails)
                {
                    outlookresponse.responsecode = "11";
                    outlookresponse.responsemessage = "All teams have sent their emails";
                }
            }
            catch (Exception ex)
            {
                LogError.Errhandler(ex.Message);
            }
            return outlookresponse;
        }

        
    }
}
