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

namespace OutlookDomainMailOrganizer
{
    public partial class ThisAddIn
    {
        #region Private Members

        Dictionary<string, Folder> domainsDb;

        #endregion

        #region Constants

        const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
        const string PR_SENT_REPRESENTING_ENTRYID = @"http://schemas.microsoft.com/mapi/proptag/0x00410102";

        #endregion

        #region Main Methods

        private void ODMOAddIn_Startup(object sender, System.EventArgs e)
        {
            InitializeDomainsDatabase();

            ProcessUnreadMessages();

            this.Application.NewMail += new ApplicationEvents_11_NewMailEventHandler(ProcessUnreadMessages);
        }

        private void ODMOAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #endregion

        #region Private Methods

        private void InitializeDomainsDatabase()
        {
            domainsDb = new Dictionary<string, Folder>();

            var inbox = (MAPIFolder)this.Application.ActiveExplorer().Session.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
            var customers = inbox.Parent.Folders["Customers"].Folders;

            foreach (object customer in customers)
            {
                var customerFolder = customer as Outlook.Folder;
                var customerName = customerFolder.Name;

                domainsDb.Add(customerName, customerFolder);
            }
        }

        private void ProcessUnreadMessages()
        {
            var inbox = (MAPIFolder)this.Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            //var inbox = (MAPIFolder)this.Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail);
            var unreadItems = (Items)inbox.Items.Restrict("[Unread]=true");
            //var unreadItems = (Items)inbox.Items;

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
                            var matchedDomain = GetDomainFromMailSender(mail) ?? GetDomainFromFirstMatchedRecipient(mail.Recipients) ?? GetDomainFromBody(mail.Body);

                            if (matchedDomain != null)
                            {
                                mail.Move(domainsDb[matchedDomain]);
                                Debug.Write("| moved <sender> (" + matchedDomain + ")");
                            }

                            break;
                        }

                    case AppointmentItem _:
                        {
                            Debug.Write(" appointment");

                            var appt = message as AppointmentItem;
                            var organizer = appt.GetOrganizer();

                            var matchedDomain = GetDomainFromAddressEntry(organizer) ?? GetDomainFromFirstMatchedRecipient(appt.Recipients) ?? GetDomainFromBody(appt.Body);

                            if (matchedDomain != null)
                            {
                                appt.Move(domainsDb[matchedDomain]);
                                Debug.Write("| moved <sender> (" + matchedDomain + ")");
                            }
                            
                            break;
                        }

                    case MeetingItem _:
                        {
                            Debug.Write(" meeting");

                            var meeting = message as MeetingItem;
                            var matchedDomain = GetDomainFromMeetingOrganizer(meeting) ?? GetDomainFromFirstMatchedRecipient(meeting.Recipients) ?? GetDomainFromBody(meeting.Body);

                            if (matchedDomain != null)
                            {
                                meeting.Move(domainsDb[matchedDomain]);
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

            string smtpAddress;

            // check to see if the addressEntry object is valid
            // for some unknown reason, this is sometimes not null but invalid
            try 
            {
                if (addressEntry.AddressEntryUserType == OlAddressEntryUserType.olExchangeUserAddressEntry || addressEntry.AddressEntryUserType == OlAddressEntryUserType.olExchangeRemoteUserAddressEntry)
                {
                    var exchUser = addressEntry.GetExchangeUser();
                    if (exchUser != null) smtpAddress = exchUser.PrimarySmtpAddress;
                    else return null; // probably left the organization (unknown, MAPI not found)
                }
            } 
            catch { return null; }
            
            
            try 
            { 
                smtpAddress = addressEntry.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS); 
            }
            catch { return null; }

            return GetDomainFromEmailAddress(smtpAddress);
        }

        private string GetDomainFromEmailAddress(string smtpAddress)
        {
            if (smtpAddress == null) return null;

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
