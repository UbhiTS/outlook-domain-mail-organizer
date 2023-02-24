using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Core;
using System.Runtime.InteropServices;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;
using System.Threading.Tasks;
using System.Diagnostics;
using static System.Net.Mime.MediaTypeNames;
using System.Threading;
using OutlookDomainMailOrganizer;
using Microsoft.Office.Tools.Ribbon;

namespace OutlookDomainMailOrganizer
{
    public partial class ThisAddIn
    {
        #region Configuration Parameters

        string domainRootFolderName = "Customers";

        #endregion

        #region Private Members

        Dictionary<string, Folder> domainsDb;
        Folder inboxFolder = null;
        Folder domainsFolder = null;
        int i;

        #endregion

        #region Properties

        bool ChronoSortEnabled
        {
            get { return Globals.Ribbons.Ribbon1.chkChronoSort.Checked; }
            set { Globals.Ribbons.Ribbon1.chkChronoSort.Checked = value; }
        }

        #endregion

        #region Constants

        const string PR_SMTP_ADDRESS                = @"http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
        const string PR_SENT_REPRESENTING_ENTRYID   = @"http://schemas.microsoft.com/mapi/proptag/0x00410102";
        const string PR_SORT_POSITION               = @"http://schemas.microsoft.com/mapi/proptag/0x30200102";

        #endregion

        #region Main Methods

        private void ODMOAddIn_Startup(object sender, System.EventArgs e)
        {
            InitializeAddIn();
            InitializeDomainsDatabase();
            ProcessUnreadMessages(inboxFolder);
            SortFoldersByChronology();
            SubscribeToEvents();

            
        }

        private void ODMOAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #endregion

        #region Event Handlers

        private void SubscribeToEvents()
        {
            Application.NewMail += NewMail;
            Globals.Ribbons.Ribbon1.btnOrganizeInbox.Click += btnOrganizeInbox_Click;
            Globals.Ribbons.Ribbon1.chkChronoSort.Click += chkChronoSort_Click;
        }

        private void NewMail()
        {
            ProcessUnreadMessages(inboxFolder);
        }

        private void btnOrganizeInbox_Click(object sender, RibbonControlEventArgs e)
        {
            ProcessUnreadMessages(inboxFolder);
        }

        private void chkChronoSort_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.ChronoSort = ChronoSortEnabled;
            Properties.Settings.Default.Save();
        }

        #endregion

        #region Private Methods

        private void InitializeAddIn()
        {
            ChronoSortEnabled = Properties.Settings.Default.ChronoSort;

            inboxFolder = (Folder)Application.ActiveExplorer().Session.GetDefaultFolder(OlDefaultFolders.olFolderInbox);

            bool domainRootExists = false;

            foreach (Folder folder in inboxFolder.Parent.Folders)
            {
                if (folder.Name == domainRootFolderName)
                {
                    domainRootExists = true;
                    break;
                }
            }

            if (domainRootExists == false) inboxFolder.Parent.Folders.Add(domainRootFolderName);

            domainsFolder = inboxFolder.Parent.Folders[domainRootFolderName];
        }

        private void InitializeDomainsDatabase()
        {
            domainsDb = new Dictionary<string, Folder>();
            
            foreach (dynamic folder in domainsFolder.Folders)
            {
                domainsDb.Add(folder.Name, folder);
            }
        }

        private void ProcessUnreadMessages(Folder folder)
        {
            var unreadItems = folder.Items.Restrict("[Unread]=true");
            var unreadItemsCount = unreadItems.Count;

            for (int i = unreadItemsCount; i > 0; i--)
            {
                Debug.Write(i.ToString());

                object message = unreadItems[i];

                if (message == null) continue;

                switch (message)
                {
                    case MailItem _:
                        {
                            Debug.Write(" mail");

                            var mail = message as MailItem;
                            var matchedDomain = GetDomainFromMailSender(mail) 
                                ?? GetDomainFromFirstMatchedRecipient(mail.Recipients) 
                                ?? GetDomainFromBody(mail.Body);

                            if (matchedDomain != null)
                            {
                                var matchedFolder = domainsDb[matchedDomain];
                                mail.Move(matchedFolder);
                                MoveFolderToTop(matchedFolder);
                                Debug.Write("| moved <sender> (" + matchedDomain + ")");
                            }

                            break;
                        }

                    case AppointmentItem _:
                        {
                            Debug.Write(" appointment");

                            var appt = message as AppointmentItem;
                            var organizer = appt.GetOrganizer();

                            var matchedDomain = GetDomainFromAddressEntry(organizer) 
                                ?? GetDomainFromFirstMatchedRecipient(appt.Recipients) 
                                ?? GetDomainFromBody(appt.Body);

                            if (matchedDomain != null)
                            {
                                var matchedFolder = domainsDb[matchedDomain];
                                appt.Move(matchedFolder);
                                MoveFolderToTop(matchedFolder);
                                Debug.Write("| moved <sender> (" + matchedDomain + ")");
                            }
                            
                            break;
                        }

                    case MeetingItem _:
                        {
                            Debug.Write(" meeting");

                            var meeting = message as MeetingItem;
                            var matchedDomain = GetDomainFromMeetingOrganizer(meeting) 
                                ?? GetDomainFromFirstMatchedRecipient(meeting.Recipients) 
                                ?? GetDomainFromBody(meeting.Body);

                            if (matchedDomain != null)
                            {
                                var matchedFolder = domainsDb[matchedDomain];
                                meeting.Move(matchedFolder);
                                MoveFolderToTop(matchedFolder);
                                Debug.Write("| moved <sender> (" + matchedDomain + ")");
                            }

                            break;
                        }

                    case ContactItem _:
                        Debug.Write(" Contact");
                        break;
                    case Folder _:
                        Debug.Write(" Folder");
                        break;
                    case NoteItem _:
                        Debug.Write(" Note");
                        break;
                    case PostItem _:
                        Debug.Write(" Post");
                        break;
                    case TaskItem _:
                        Debug.Write(" Task");
                        break;
                }

                Debug.WriteLine("");
            }
        }

        private void SortFoldersByChronology()
        {
            if (!ChronoSortEnabled) return;
            
            var chronoDb = new SortedDictionary<DateTime, List<Folder>>();

            foreach (var folderName in domainsDb.Keys)
            {
                var items = domainsDb[folderName].Items;

                if (items != null)
                {
                    items.Sort("[ReceivedTime]");
                    var lastItem = items.GetLast();

                    if (lastItem != null)
                    {
                        if (chronoDb.ContainsKey(lastItem.ReceivedTime)) chronoDb[lastItem.ReceivedTime].Add(domainsDb[folderName]);
                        else chronoDb.Add(lastItem.ReceivedTime, new List<Folder>() { domainsDb[folderName] });
                        continue;
                    }
                }
                
                if (chronoDb.ContainsKey(DateTime.MinValue)) chronoDb[DateTime.MinValue].Add(domainsDb[folderName]);
                else chronoDb.Add(DateTime.MinValue, new List<Folder>() { domainsDb[folderName] });
            }

            i = 255;

            foreach (var folders in chronoDb)
            {
                foreach (var folder in folders.Value)
                {
                    var currentPosition = folder.PropertyAccessor.BinaryToString(folder.PropertyAccessor.GetProperty(PR_SORT_POSITION));
                    var newPosition = i.ToString("X2");

                    if (currentPosition != newPosition)
                    {
                        folder.PropertyAccessor.SetProperty(PR_SORT_POSITION, folder.PropertyAccessor.StringToBinary(newPosition));
                    }

                    i--;
                }
            }
        }

        private void MoveFolderToTop(Folder folder)
        {
            var currentPosition = folder.PropertyAccessor.BinaryToString(folder.PropertyAccessor.GetProperty(PR_SORT_POSITION));
            var topPosition = (i + 1).ToString("X2");

            if (currentPosition != topPosition)
            {
                folder.PropertyAccessor.SetProperty(PR_SORT_POSITION, folder.PropertyAccessor.StringToBinary(i.ToString("X2")));
                i--;
            }

            if (i == 0) { SortFoldersByChronology(); }
        }

        #endregion

        #region Helper Methods

        private string GetDomainFromMailSender(MailItem mail)
        {
            // https://learn.microsoft.com/en-us/office/client-developer/outlook/pia/how-to-get-the-smtp-address-of-the-sender-of-a-mail-item

            if (mail == null) return null;

            if (mail.SenderEmailType == "EX") return GetDomainFromAddressEntry(mail.Sender);

            return GetDomainFromEmailAddress(mail.SenderEmailAddress);
        }

        private string GetDomainFromMeetingOrganizer(MeetingItem meeting)
        {
            if (meeting == null) return null;

            var organizerEntryID = meeting.PropertyAccessor.BinaryToString(meeting.PropertyAccessor.GetProperty(PR_SENT_REPRESENTING_ENTRYID));
            var organizer = Application.Session.GetAddressEntryFromID(organizerEntryID);

            if (organizer == null) return null;

            return GetDomainFromAddressEntry(organizer);
        }

        private string GetDomainFromFirstMatchedRecipient(Recipients recipients)
        {
            if (recipients == null) return null;
            
            foreach (Recipient recipient in recipients)
            {
                var domain = GetDomainFromEmailAddress(GetEmailAddressFromRecipient(recipient));

                if (domain != null) return domain;
            }

            return null;
        }

        private string GetEmailAddressFromRecipient(Recipient recipient)
        {
            if (recipient == null) return null;
            
            // https://learn.microsoft.com/en-us/office/client-developer/outlook/pia/how-to-get-the-e-mail-address-of-a-recipient

            var pa = recipient.PropertyAccessor;
            string smtpAddress = null;

            try { smtpAddress = pa.GetProperty(PR_SMTP_ADDRESS); }
            catch (System.Exception ex) { Debug.Write(ex.Message); } // probably left the organization (unknown)

            return smtpAddress;
        }

        private string GetDomainFromAddressEntry(AddressEntry addressEntry)
        {
            if (addressEntry == null) return null;

            string smtpAddress = null;

            // check to see if the addressEntry object is valid
            // for some unknown reason, this is sometimes not null but invalid
            try 
            {
                if (addressEntry.AddressEntryUserType == OlAddressEntryUserType.olExchangeUserAddressEntry || addressEntry.AddressEntryUserType == OlAddressEntryUserType.olExchangeRemoteUserAddressEntry)
                {
                    var exchUser = addressEntry.GetExchangeUser();

                    if (exchUser != null)
                    {
                        smtpAddress = exchUser.PrimarySmtpAddress;
                    }
                    else
                    {
                        // else the user has probably left the organization (unknown, MAPI not found)
                    }
                }
            } 
            catch { }
            
            try 
            {
                if (smtpAddress == null || smtpAddress == string.Empty)
                {
                    smtpAddress = addressEntry.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS);
                }
            }
            catch { }

            try
            {
                if (smtpAddress == null || smtpAddress == string.Empty)
                {
                    smtpAddress = addressEntry.Address;
                }
            }
            catch { }

            return GetDomainFromEmailAddress(smtpAddress);
        }

        private string GetDomainFromEmailAddress(string smtpAddress)
        {
            if (smtpAddress == null || smtpAddress == string.Empty) return null;

            // Debug.Write(" " + smtpAddress);

            var smtpAddressArray = smtpAddress.Split('@');

            if (smtpAddressArray.Length == 2)
            {
                var domain = smtpAddressArray[1];

                if (domainsDb.ContainsKey(domain))
                {
                    return domain;
                }
            }

            return null;
        }

        private string GetDomainFromBody(string body)
        {
            if (body == null || body.Length <= 4) return null;

            foreach (var domain in domainsDb.Keys)
            {
                if (body.Contains(domain))
                {
                    return domain;
                }
            }

            return null;
        }

        #endregion

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ODMOAddIn_Startup);
            this.Shutdown += new System.EventHandler(ODMOAddIn_Shutdown);
        }

        #endregion
    }
}
