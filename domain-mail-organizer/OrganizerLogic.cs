using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Xml.Linq;
using Microsoft.Office.Interop.Outlook;

namespace DomainMailOrganizer
{
    public class OrganizerLogic
    {
        #region Events

        public delegate void MessagesRemainingEvent(int messagesRemaining);
        public event MessagesRemainingEvent MessagesRemainingEventHandler;

        #endregion

        #region Constants

        const string PR_SMTP_ADDRESS = @"http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
        const string PR_SENT_REPRESENTING_ENTRYID = @"http://schemas.microsoft.com/mapi/proptag/0x00410102";
        const string PR_SORT_POSITION = @"http://schemas.microsoft.com/mapi/proptag/0x30200102";

        const string ARCHIVE_FOLDER_NAME = @"Archive";

        #endregion

        #region Private Members

        Application outlook = null;
        Folder inboxFolder = null;
        Folder archiveFolder = null;
        Folder domainsFolder = null;

        Dictionary<string, Folder> domainsDb;
        Dictionary<string, Folder> keywordsDb;

        bool dbsInit = false;

        bool sortEnabled = true;
        bool foldersSortInit = false;
        int foldersSortCtr = 0;

        #endregion

        #region Constructor

        public OrganizerLogic(Application app, string domainsFolderName, bool chronoSortEnabled)
        {
            outlook = app;
            domainsFolder = GetOutlookFolder(domainsFolderName);
            archiveFolder = GetOutlookFolder(ARCHIVE_FOLDER_NAME);
            sortEnabled = chronoSortEnabled;
        }

        #endregion

        #region Public Methods

        public void ProcessInbox1Day()
        {
            var filter = "[ReceivedTime] > '" + DateTime.Now.AddDays(-1).ToString("MM/dd/yyyy HH:mm") + "'";
            ProcessMessages(inboxFolder, filter);
        }

        public void ProcessInbox2Day()
        {
            var filter = "[ReceivedTime] > '" + DateTime.Now.AddDays(-2).ToString("MM/dd/yyyy HH:mm") + "'";
            ProcessMessages(inboxFolder, filter);
        }

        public void ProcessInbox7Day()
        {
            var filter = "[ReceivedTime] > '" + DateTime.Now.AddDays(-7).ToString("MM/dd/yyyy HH:mm") + "'";
            ProcessMessages(inboxFolder, filter);
        }

        public void ProcessInbox30Day()
        {
            var filter = "[ReceivedTime] > '" + DateTime.Now.AddDays(-30).ToString("MM/dd/yyyy HH:mm") + "'";
            ProcessMessages(inboxFolder, filter);
        }

        public void ProcessInboxAll()
        {
            ProcessMessages(inboxFolder, null);
        }

        public void ProcessArchive1Day()
        {
            var filter = "[ReceivedTime] > '" + DateTime.Now.AddDays(-1).ToString("MM/dd/yyyy HH:mm") + "'";
            ProcessMessages(archiveFolder, filter);
        }

        public void ProcessArchive2Day()
        {
            var filter = "[ReceivedTime] > '" + DateTime.Now.AddDays(-2).ToString("MM/dd/yyyy HH:mm") + "'";
            ProcessMessages(archiveFolder, filter);
        }

        public void ProcessArchive7Day()
        {
            var filter = "[ReceivedTime] > '" + DateTime.Now.AddDays(-7).ToString("MM/dd/yyyy HH:mm") + "'";
            ProcessMessages(archiveFolder, filter);
        }

        public void ProcessArchive30Day()
        {
            var filter = "[ReceivedTime] > '" + DateTime.Now.AddDays(-30).ToString("MM/dd/yyyy HH:mm") + "'";
            ProcessMessages(archiveFolder, filter);
        }

        public void ProcessArchiveAll()
        {
            ProcessMessages(archiveFolder, null);
        }

        public void ArchiveAllInboxItems()
        {
            ArchiveInboxItems();
        }

        public void ChronoSortFolders()
        {
            if (!sortEnabled) return;

            if (dbsInit == false) InitDBs();

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

            foldersSortCtr = 255;

            foreach (var folders in chronoDb)
            {
                foreach (var folder in folders.Value)
                {
                    string currentPosition = null;

                    try { currentPosition = folder.PropertyAccessor.BinaryToString(folder.PropertyAccessor.GetProperty(PR_SORT_POSITION)); } catch { }

                    var newPosition = foldersSortCtr.ToString("X2");

                    if (currentPosition == null || currentPosition != newPosition)
                    {
                        folder.PropertyAccessor.SetProperty(PR_SORT_POSITION, folder.PropertyAccessor.StringToBinary(newPosition));
                    }

                    foldersSortCtr--;
                }
            }

            foldersSortInit = true;
        }

        #endregion

        #region Private Methods

        private void InitDBs()
        {
            domainsDb = new Dictionary<string, Folder>();
            keywordsDb = new Dictionary<string, Folder>();

            foreach (Folder folder in domainsFolder.Folders)
            {
                domainsDb.Add(folder.Name.ToLower(), folder);

                if (folder.Description != null && folder.Description != string.Empty)
                {
                    var keywords = folder.Description.ToLower()
                        .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(keyword => keyword.Trim())
                        .ToArray();


                    foreach (var keyword in keywords)
                    {
                        if (!keywordsDb.ContainsKey(keyword))
                        {
                            keywordsDb.Add(keyword.Trim(), folder);
                        }
                        else
                        {
                            // TODO: update error status on ribbon for duplicate keyword found; 
                        }
                    }
                }
            }

            dbsInit = true;
        }

        private Folder GetOutlookFolder(string folderName)
        {
            inboxFolder = (Folder)outlook.ActiveExplorer().Session.GetDefaultFolder(OlDefaultFolders.olFolderInbox);

            bool folderExists = false;

            foreach (Folder folder in inboxFolder.Parent.Folders)
            {
                if (folder.Name == folderName)
                {
                    folderExists = true;
                    break;
                }
            }

            if (folderExists == false) inboxFolder.Parent.Folders.Add(folderName);

            return inboxFolder.Parent.Folders[folderName];
        }

        private void ProcessMessages(Folder folder, string filter)
        {
            if (dbsInit == false) InitDBs();
            if (foldersSortInit == false) ChronoSortFolders();

            Items items;

            if (filter == null || filter == string.Empty)
            {
                items = folder.Items;
            }
            else
            {
                items = folder.Items.Restrict(filter);
            }
            
            var unreadItemsCount = items.Count;

            for (int i = unreadItemsCount; i > 0; i--)
            {
                Debug.Write(i.ToString());

                object message = items[i];

                if (message == null) continue;

                switch (message)
                {
                    case MailItem _:
                        {
                            Debug.Write(" mail");

                            var mail = message as MailItem;
                            var matchedDomain = GetDomainFromMailSender(mail)
                                ?? GetDomainFromFirstMatchedRecipient(mail.Recipients)
                                ?? GetDomainFromSubject(mail.Subject)
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
                                ?? GetDomainFromSubject(appt.Subject)
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
                                ?? GetDomainFromSubject(meeting.Subject)
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
                    case ReportItem _:
                        {
                            Debug.Write(" report");

                            var report = message as ReportItem;
                            var matchedDomain = GetDomainFromSubject(report.Subject)
                                ?? GetDomainFromBody(report.Body);

                            if (matchedDomain != null)
                            {
                                var matchedFolder = domainsDb[matchedDomain];
                                report.Move(matchedFolder);
                                MoveFolderToTop(matchedFolder);
                                Debug.Write("| moved <sender> (" + matchedDomain + ")");
                            }

                        }
                        break;
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
                    default:
                        Debug.Write(" Other");
                        break;
                }

                Debug.WriteLine("");

                MessagesRemainingEventHandler?.Invoke(i-1);
            }
        }

        private void MoveFolderToTop(Folder folder)
        {
            if (!sortEnabled) return;

            var currentPosition = folder.PropertyAccessor.BinaryToString(folder.PropertyAccessor.GetProperty(PR_SORT_POSITION));
            var topPosition = (foldersSortCtr + 1).ToString("X2");

            if (currentPosition != topPosition)
            {
                folder.PropertyAccessor.SetProperty(PR_SORT_POSITION, folder.PropertyAccessor.StringToBinary(foldersSortCtr.ToString("X2")));
                foldersSortCtr--;
            }

            if (foldersSortCtr == 0)
            {
                ChronoSortFolders();
                //MessageBox.Show("Please restart Outlook to restore the chronological folder sorting.", "Add-In: Domain Mail Organizer", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                Debug.WriteLine("Please restart Outlook to restore the chronological folder sorting.");
            }
        }

        private void ArchiveInboxItems()
        {
            var items = inboxFolder.Items;
            var itemsCount = items.Count;

            for (int i = itemsCount; i > 0; i--)
            {
                Debug.Write(i.ToString());

                object message = items[i];

                if (message == null) continue;

                switch (message)
                {
                    case MailItem _:
                        {
                            Debug.Write(" mail");
                            var mail = message as MailItem;
                            mail.Move(archiveFolder);
                            break;
                        }

                    case AppointmentItem _:
                        {
                            Debug.Write(" appointment");
                            var appt = message as AppointmentItem;
                            appt.Move(archiveFolder);
                            break;
                        }

                    case MeetingItem _:
                        {
                            Debug.Write(" meeting");
                            var meeting = message as MeetingItem;
                            meeting.Move(archiveFolder);
                            break;
                        }

                    case ReportItem _:
                        {
                            Debug.Write(" report");
                            var report = message as ReportItem;
                            report.Move(archiveFolder);
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
                    default:
                        Debug.Write(" Other");
                        break;
                }

                Debug.WriteLine("");

                MessagesRemainingEventHandler?.Invoke(i - 1);
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
            var organizer = outlook.Session.GetAddressEntryFromID(organizerEntryID);

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

            string smtpAddress = null;

            try { smtpAddress = recipient.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS); }
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

            var smtpAddressArray = smtpAddress.Split('@');

            if (smtpAddressArray.Length == 2)
            {
                var domain = smtpAddressArray[1].ToLower();

                if (domainsDb.ContainsKey(domain))
                {
                    return domain;
                }
            }

            return null;
        }

        private string GetDomainFromSubject(string subject)
        {
            if (subject == null || subject.Length == 0) return null;

            subject = subject.ToLower();

            foreach (var domain in domainsDb.Keys)
            {
                if (subject.Contains(domain))
                {
                    return domain;
                }
            }

            foreach (var keyword in keywordsDb.Keys)
            {
                if (subject.Contains(keyword))
                {
                    return keywordsDb[keyword].Name;
                }
            }

            return null;
        }

        private string GetDomainFromBody(string body)
        {
            if (body == null || body.Length == 0) return null;

            body = body.ToLower();

            foreach (var domain in domainsDb.Keys)
            {
                if (body.Contains(domain))
                {
                    return domain;
                }
            }

            foreach (var keyword in keywordsDb.Keys)
            {
                if (body.Contains(keyword))
                {
                    return keywordsDb[keyword].Name;
                }
            }

            return null;
        }

        #endregion
    }
}
