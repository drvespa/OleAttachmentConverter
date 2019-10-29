using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Exchange.WebServices.Data;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;


namespace Microsoft.Samples.CreateItemWithRtfBody
{
    class Program
    {
        struct CMDLINEARGS
        {
            public string Server;
            public string username;
            public string password;
            public string domain;
            public string EwsUrl;
            public bool AcceptSslCheck;
            public string TargetSubject;
        }

        public static ExchangeService g_svcObject;

        static void DisplayUsage()
        {
            Console.WriteLine("CreateItemWithRtfBody.exe <Server> [UserName] [Password] [Domain] [Subject of original email] [Accept SSL Checks] [-?]");
            Console.WriteLine("CreateItemWithRtfBody.exe ContosoServer01 jsmith T%nt0wn Contoso \"My Email Subject\" true ");
            Console.WriteLine("CreateItemWithRtfBody.exe ContosoServer01 \"\" \"\" \"\" \"My Email Subject\"");
            Console.WriteLine("CreateItemWithRtfBody.exe -?");
            Console.WriteLine("CreateItemWithRtfBody.exe -help");
        }

        static void Main(string[] args)
        {
            var cmdLineArgs = new CMDLINEARGS();

            if (args.Length == 0 || args[0][1] == '?' || args[0][1] == 'H' || args[0][1] == 'h')
            {
                DisplayUsage();
                return;
            }
            
            cmdLineArgs.Server = args[0];
            cmdLineArgs.EwsUrl = String.Format("https://{0}/ews/exchange.asmx", cmdLineArgs.Server);;

            g_svcObject = new ExchangeService(ExchangeVersion.Exchange2013_SP1);

            if (args.Length > 3)
            {
                if (!String.IsNullOrEmpty(args[1]))
                    cmdLineArgs.username = args[1];
                if (!String.IsNullOrEmpty(args[2]))
                    cmdLineArgs.password = args[2];
                if (!String.IsNullOrEmpty(args[3]))
                    cmdLineArgs.domain = args[3];

                if (!String.IsNullOrEmpty(cmdLineArgs.username) &&
                !String.IsNullOrEmpty(cmdLineArgs.password))
                {
                    g_svcObject.Credentials = new WebCredentials(cmdLineArgs.username,
                        cmdLineArgs.password,
                        cmdLineArgs.domain);
                }
            }
            else
            {
                DisplayUsage();
            }

            if (String.IsNullOrEmpty(cmdLineArgs.username) &&
                String.IsNullOrEmpty(cmdLineArgs.password) &&
                String.IsNullOrEmpty(cmdLineArgs.domain))
            {
                g_svcObject.UseDefaultCredentials = true;
                g_svcObject.Credentials = null;
            }

            if (args.Length > 4)
            {
                cmdLineArgs.TargetSubject = args[4];
            }

            // I guess we should do the right thing and default to false.
            cmdLineArgs.AcceptSslCheck = false;

            if (args.Length > 5)
            {
                switch (args[5][0])
                {
                    case '1':
                    case 't':
                        cmdLineArgs.AcceptSslCheck = true;
                        break;
                }
            }

            ServicePointManager.ServerCertificateValidationCallback = (object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) =>
            {
                return cmdLineArgs.AcceptSslCheck;
            };

            g_svcObject.Url = new System.Uri(cmdLineArgs.EwsUrl);
            g_svcObject.TraceEnabled = true;
            g_svcObject.TraceFlags = TraceFlags.All;

            var itemView = new ItemView(50);
            itemView.PropertySet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject);
            itemView.Traversal = ItemTraversal.Shallow;
            var searchFilter = new SearchFilter.ContainsSubstring(ItemSchema.Subject, cmdLineArgs.TargetSubject);
            var items = g_svcObject.FindItems(WellKnownFolderName.Inbox, searchFilter, itemView);
            if (items.TotalCount != 1)
            {
                if (items.TotalCount == 0)
                    Console.WriteLine("Found few items. Confirm the search terms are correct!");
                else
                    Console.WriteLine("Found too many items. Confirm the search terms are correct!");
                return;
            }

            var msg = items.Items[0];
            var rtfCompressed = new ExtendedPropertyDefinition(0x1009, MapiPropertyType.Binary);
            msg.Load(new PropertySet(BasePropertySet.FirstClassProperties, rtfCompressed, ItemSchema.Attachments));
            //var fileAttach = msg.Attachments[0];

            var newMsg = new EmailMessage(g_svcObject);
            newMsg.Subject = "RE: " + msg.Subject;
            newMsg.SetExtendedProperty(rtfCompressed, msg.ExtendedProperties[0].Value);
            if (msg.Attachments.Count > 0)
            {
                FileAttachment fileAttachment = msg.Attachments[0] as FileAttachment;
                fileAttachment.Load();
                var newAttachment = newMsg.Attachments.AddFileAttachment(fileAttachment.Name, fileAttachment.Content);
                newAttachment.IsInline = newAttachment.IsInline;
                newAttachment.IsContactPhoto = fileAttachment.IsContactPhoto;
            }
            newMsg.Save();
        }
    }
}
