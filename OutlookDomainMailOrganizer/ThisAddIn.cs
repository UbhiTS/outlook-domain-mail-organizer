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

        #endregion

        #region Main Methods

        private void ODMOAddIn_Startup(object sender, System.EventArgs e)
        {
            InitializeDomainsDatabase();

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

                switch (message)
                {
                    case MailItem _ when message != null:
                        {
                            Debug.Write(" mail");

                            var mail = message as MailItem;

                            var matchedDomain = GetMailSenderDomain(mail);

                            if (matchedDomain != null)
                            {
                                mail.Move(domainsDb[matchedDomain]);
                                Debug.Write("| moved <sender> (" + matchedDomain + ")");
                                continue;
                            }
                            else
                            {
                                matchedDomain = GetFirstMatchedRecipientDomain(mail.Recipients);

                                if (matchedDomain != null)
                                {
                                    mail.Move(domainsDb[matchedDomain]);
                                    Debug.Write("| moved <recipient> (" + matchedDomain + ")");
                                }
                            }

                            break;
                        }

                    case AppointmentItem _:
                        {
                            Debug.Write(" appointment");

                            var appointment = message as AppointmentItem;
                            var matchedDomain = GetFirstMatchedRecipientDomain(appointment.Recipients);
                            
                            if (matchedDomain != null)
                            {
                                appointment.Move(domainsDb[matchedDomain]);
                                Debug.Write("| moved <recipient> (" + matchedDomain + ")");
                            }

                            break;
                        }

                    case MeetingItem _:
                        {
                            Debug.Write(" meeting");

                            var meeting = message as MeetingItem;
                            var matchedDomain = GetFirstMatchedRecipientDomain(meeting.Recipients);

                            if (matchedDomain != null)
                            {
                                meeting.Move(domainsDb[matchedDomain]);
                                Debug.Write("| moved <recipient> (" + matchedDomain + ")");
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

        private string GetMailSenderDomain(Outlook.MailItem mail)
        {
            // https://learn.microsoft.com/en-us/office/client-developer/outlook/pia/how-to-get-the-smtp-address-of-the-sender-of-a-mail-item

            if (mail == null) return null;

            string senderSmtpAddress = null;

            if (mail.SenderEmailType == "EX")
            {
                var sender = mail.Sender;

                if (sender == null) return null;

                if (sender.AddressEntryUserType == OlAddressEntryUserType.olExchangeUserAddressEntry || sender.AddressEntryUserType == OlAddressEntryUserType.olExchangeRemoteUserAddressEntry)
                {
                    var exchUser = sender.GetExchangeUser();
                    if (exchUser != null) senderSmtpAddress = exchUser.PrimarySmtpAddress;
                    else return null; // probably left the organization (unknown, MAPI not found)
                }
                else
                {
                    try { senderSmtpAddress = sender.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS); }
                    catch (System.Exception ex) { Debug.Write(ex.Message); }
                }
            }
            else
            {
                senderSmtpAddress = mail.SenderEmailAddress;
            }

            if (senderSmtpAddress != null)
            {
                Debug.Write(" " + senderSmtpAddress);

                var senderSmtpAddressArray = senderSmtpAddress.Split('@');
                if (senderSmtpAddressArray.Length > 2)
                {
                    var senderDomain = senderSmtpAddressArray[1];

                    if (domainsDb.ContainsKey(senderDomain))
                    {
                        return senderDomain;
                    }
                }
            }

            return null;
        }

        private string GetFirstMatchedRecipientDomain(Recipients recipients)
        {
            foreach (Recipient recipient in recipients)
            {
                var recipientSmtpAddress = GetSMTPAddressForRecipient(recipient);

                if (recipientSmtpAddress != null)
                {
                    Debug.Write(" " + recipientSmtpAddress);

                    var recipientSmtpAddressArray = recipientSmtpAddress.Split('@');

                    if (recipientSmtpAddressArray.Length < 2) continue;

                    var recipientDomain = recipientSmtpAddressArray[1];

                    if (domainsDb.ContainsKey(recipientDomain))
                    {
                        return recipientDomain;
                    }
                }
            }

            return null;
        }

        private string GetSMTPAddressForRecipient(Recipient recipient)
        {
            // https://learn.microsoft.com/en-us/office/client-developer/outlook/pia/how-to-get-the-e-mail-address-of-a-recipient

            var pa = recipient.PropertyAccessor;
            string smtpAddress = null;

            try { smtpAddress = pa.GetProperty(PR_SMTP_ADDRESS); }
            catch (System.Exception ex) { Debug.Write(ex.Message); } // probably left the organization (unknown)

            return smtpAddress;
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
