using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Configuration;
using Microsoft.Exchange.WebServices.Data;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.IO;



namespace AttachmentsDownloader
{
    public partial class AttachmentsDownloader : ServiceBase
    {
        private string sLog;
        private string sSource;
        private string sEvent;
        private Thread Scanner;
         
        public AttachmentsDownloader()
        {
            sLog         = "Application";
            sSource      = "AttachmentsDownloader";
            InitializeComponent();
        }

              
        public void OnDebug()
        {
           OnStartAfter();
        }

        protected override void OnStart(string[] args)
        {
            Scanner = new Thread(new ThreadStart(OnStartAfter));
            Scanner.Start();
        }
        protected override void OnStop()
        {
            Scanner.Abort();

        }
        
        private void OnStartAfter()
        {
            string sMailBox = System.String.Empty;
            string sMailServerUrl = System.String.Empty;
            string sUser = System.String.Empty;
            string sPassword = System.String.Empty;
            string sDomain = System.String.Empty;
            string sFolderRoot = System.String.Empty;
            int iItemsProccesing = 0;
            int iScanDelay = 0;
            Boolean bFolderDestDomain = false;
            Boolean bDestinationFileFormatLong = false;
            Boolean bProcessingFileToLog = false;

            if (!EventLog.SourceExists(sSource))
                EventLog.CreateEventSource(sSource, sLog);
            sEvent = "Service started";

            EventLog.WriteEntry(sSource, sEvent);


            try
            {
                sMailBox = ConfigurationManager.AppSettings["MailBox"];
                sMailServerUrl = ConfigurationManager.AppSettings["MailServerUrl"];
                sUser = ConfigurationManager.AppSettings["User"];
                sPassword = ConfigurationManager.AppSettings["Password"];
                sDomain = ConfigurationManager.AppSettings["Domain"];
                sFolderRoot = ConfigurationManager.AppSettings["FolderRoot"];
                iItemsProccesing = Convert.ToInt32(ConfigurationManager.AppSettings["ItemsProccesing"]);        // How much emails retrieve to view
                iScanDelay = Convert.ToInt32(ConfigurationManager.AppSettings["ScanDelay"]);                    // Delay between scan Inbox
                bFolderDestDomain = Convert.ToBoolean(ConfigurationManager.AppSettings["FolderDestinationNameWithDomain"]);
                bDestinationFileFormatLong = Convert.ToBoolean(ConfigurationManager.AppSettings["DestinationFileFormatLong"]); //
                bProcessingFileToLog = Convert.ToBoolean(ConfigurationManager.AppSettings["ProcessingFileToLog"]); //
                
                sEvent =    "MailBox=" + sMailBox + "\n" +
                            "MailServerUrl=" + sMailServerUrl + "\n" +
                            "User=" + sUser + "\n" +
                            "Domain=" + sDomain + "\n" +
                            "FolderRoot=" + sFolderRoot + "\n" +
                            "ItemsProccesing=" + iItemsProccesing + "\n" +
                            "ScanDelay=" + iScanDelay + "\n" +
                            "FolderDestDomain=" + bFolderDestDomain + "\n" +
                            "DestinationFileFormatLong=" + bDestinationFileFormatLong + "\n" +
                            "ProcessingFileToLog=" + bProcessingFileToLog;

                EventLog.WriteEntry(sSource, sEvent);
           
           
            }
            catch (Exception e)
            {
                EventLog.WriteEntry(sSource, e.Message, EventLogEntryType.Warning, 234);
            }

            ////////////////////;
            //
            Boolean bItemDelete = false;
            string sFilePath = System.String.Empty;



            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP1);
            ServicePointManager.ServerCertificateValidationCallback = (obj, certificate, chain, errors) => true;

            service.Credentials = new System.Net.NetworkCredential(sUser, sPassword, sDomain);

            Uri uriMailServerUrl = new Uri(sMailServerUrl);
            service.Url = uriMailServerUrl;

            FolderId InboxId = new FolderId(WellKnownFolderName.Inbox, sMailBox);
            FindItemsResults<Item> findResults = null;

            while (true)
            {

                try
                {

                    findResults = service.FindItems(InboxId, new ItemView(iItemsProccesing));
                }

                catch (Exception e)
                {
                    sEvent = "Exchange Server:" +e.Source + ":" + e.Message;
                    EventLog.WriteEntry(sSource, sEvent, EventLogEntryType.Warning, 234);                  
                }

                if (findResults != null)
                {

                    foreach (Item message in findResults.Items)
                    {
                        bItemDelete = true;
                        if (message.HasAttachments)
                        {
                            EmailMessage emailMessage = EmailMessage.Bind(service, message.Id, new PropertySet(BasePropertySet.IdOnly, ItemSchema.Attachments));
#if DEBUG
                            sEvent  = "Debug event Subject: " + message.Subject + " Date: " + message.DateTimeReceived;
                            EventLog.WriteEntry(sSource, sEvent );
#endif

                            
                            
                            emailMessage.Load();
                            string sSenderAdddress = emailMessage.From.Address;
                            String sFolderPath = GetPath(bFolderDestDomain, message.DateTimeReceived, sFolderRoot, sSenderAdddress);

                            foreach (Attachment attachment in emailMessage.Attachments)
                            {
                                FileAttachment fileAttachment = attachment as FileAttachment;

                                if (bDestinationFileFormatLong)
                                {
                                    string sNewFileName = AttachFileName(fileAttachment.Name, message.DateTimeReceived);
                                    sFilePath = sFolderPath + sNewFileName;
                                }
                                else
                                {
                                    sFilePath = sFolderPath + fileAttachment.Name;
                                }

                                if (bProcessingFileToLog)
                                {
                                    sEvent = "File stored path=" + sFilePath;
                                    EventLog.WriteEntry(sSource, sEvent);
                                }

                                try
                                {
                                    fileAttachment.Load(sFilePath);

                                }

                                    
                                catch (System.UnauthorizedAccessException e)
                                {
                                    sEvent = e.Source + ":" + e.Message;
                                    EventLog.WriteEntry(sSource, sEvent, EventLogEntryType.Warning, 234);
                                    bItemDelete = false;
                                }
                                catch (Exception e)
                                {
                                    sEvent = e.Source + ":" + e.Message;
                                    EventLog.WriteEntry(sSource, sEvent, EventLogEntryType.Warning, 234);
                                    bItemDelete = false;
                                }
                            }


                        }

                        //  Move processed item to \Deleted items
                        if (bItemDelete) message.Delete(DeleteMode.MoveToDeletedItems);

                    }
                }

                System.Threading.Thread.Sleep(iScanDelay);
            }
           

        }

        private string GetPath(Boolean abFolderDestDomain, DateTime adtDateTimeReceived, string asFolderRoot, string asSenderAddress)
        {
            string sFolderDestination = System.String.Empty;
            string sFolderDestinationFull = System.String.Empty;

            // folder format \YYYY\MMYY\DDMMYYYY\
            sFolderDestination = asFolderRoot + string.Format("{0:yyyy}", adtDateTimeReceived) + "\\" + string.Format("{0:MMyy}", adtDateTimeReceived) + "\\" + string.Format("{0:ddMMyy}", adtDateTimeReceived) + "\\";


            if (abFolderDestDomain)
            {
                // folder format \YYYY\MMYY\DDMMYYYY\SenderEmailDomain
                string sEmailDomain = asSenderAddress.Substring(asSenderAddress.LastIndexOf("@") + 1, asSenderAddress.Length - asSenderAddress.LastIndexOf("@") - 1);
                sFolderDestinationFull = sFolderDestination + sEmailDomain + "\\";
            }
            else
            {
                sFolderDestinationFull = sFolderDestination;
            }



            try
            {
                if (!Directory.Exists(sFolderDestinationFull))
                {
                    Directory.CreateDirectory(sFolderDestinationFull);
                }
            }
            catch (IOException e)
            {
               EventLog.WriteEntry(sSource, e.Message, EventLogEntryType.Warning, 234);
            }

            return sFolderDestinationFull;
        }
        private string AttachFileName(string asFileName, DateTime adtDateTimeReceived)
        {
            string sFileExtension = asFileName.Substring(asFileName.LastIndexOf(".") + 1, 3);
            string sFileName = asFileName.Substring(0, asFileName.LastIndexOf("."));
            string sFileNameAddition = string.Format("{0:ddMMyyyy}", adtDateTimeReceived) + "_" + string.Format("{0:HHmmss}", adtDateTimeReceived);
            string sNewFileName = sFileName + "_" + sFileNameAddition + "." + sFileExtension;

            return sNewFileName;
        }

    }
}
