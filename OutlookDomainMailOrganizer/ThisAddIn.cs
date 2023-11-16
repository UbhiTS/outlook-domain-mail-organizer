﻿using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Outlook;
using System;

namespace OutlookDomainMailOrganizer
{
    public partial class ThisAddIn
    {
        #region Configuration Parameters

        string inboxArchiveFolderName = "Inbox-Archive";
        string domainsFolderName = "Customers";

        private readonly object _lockObject = new object();
        System.Threading.Thread t = null;

        #endregion

        #region Properties

        DomainMailOrganizer.OrganizerLogic organizerLogic = null;

        #endregion

        #region Main Methods

        private void ODMOAddIn_Startup(object sender, System.EventArgs e)
        {
            Globals.Ribbons.Ribbon1.chkChronoSort.Checked = true;

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
            Globals.Ribbons.Ribbon1.btnRefresh.Click += btnRefresh_Click;
            Globals.Ribbons.Ribbon1.chkChronoSort.Click += chkChronoSort_Click;
            Globals.Ribbons.Ribbon1.btnOrganizeInbox.Click += btnOrganizeInbox_Click;
            Globals.Ribbons.Ribbon1.btnOrganizeArchive.Click += btnOrganizeArchive_Click;

            Application.NewMail += NewMail;
        }

        private void btnRefresh_Click(object sender, RibbonControlEventArgs e)
        {
            if (t != null && t.ThreadState == System.Threading.ThreadState.Running) return;

            organizerLogic = null;
        }

        private void chkChronoSort_Click(object sender, RibbonControlEventArgs e)
        {
            if (t != null && t.ThreadState == System.Threading.ThreadState.Running) return;

            organizerLogic = null;
        }

        private void btnOrganizeInbox_Click(object sender, RibbonControlEventArgs e)
        {
            if (t != null && t.ThreadState == System.Threading.ThreadState.Running) return;

            InitializeOrganizerLogic();

            switch (int.Parse(Globals.Ribbons.Ribbon1.ddDays.SelectedItem.Tag.ToString()))
            {
                case 1:
                    t = new System.Threading.Thread(organizerLogic.ProcessInbox1Day);
                    break;
                case 7:
                    t = new System.Threading.Thread(organizerLogic.ProcessInbox7Day);
                    break;
                case 30:
                    t = new System.Threading.Thread(organizerLogic.ProcessInbox30Day);
                    break;
                default:
                    t = new System.Threading.Thread(organizerLogic.ProcessInboxAll);
                    break;
            }

            t.SetApartmentState(System.Threading.ApartmentState.STA);
            t.Start();
        }

        private void btnOrganizeArchive_Click(object sender, RibbonControlEventArgs e)
        {
            if (t != null && t.ThreadState == System.Threading.ThreadState.Running) return;

            InitializeOrganizerLogic();

            switch (int.Parse(Globals.Ribbons.Ribbon1.ddDays.SelectedItem.Tag.ToString()))
            {
                case 1:
                    t = new System.Threading.Thread(organizerLogic.ProcessArchive1Day);
                    break;
                case 7:
                    t = new System.Threading.Thread(organizerLogic.ProcessArchive7Day);
                    break;
                case 30:
                    t = new System.Threading.Thread(organizerLogic.ProcessArchive30Day);
                    break;
                default:
                    t = new System.Threading.Thread(organizerLogic.ProcessArchiveAll);
                    break;
            }

            t.SetApartmentState(System.Threading.ApartmentState.STA);
            t.Start();
        }

        private void NewMail()
        {
            if (t != null && t.ThreadState == System.Threading.ThreadState.Running) return;

            InitializeOrganizerLogic();

            t = new System.Threading.Thread(organizerLogic.ProcessInboxUnread);
            t.SetApartmentState(System.Threading.ApartmentState.STA);
            t.Start();
        }

        private void InitializeOrganizerLogic()
        {
            lock (_lockObject)
            {
                if (organizerLogic == null)
                {
                    organizerLogic = new DomainMailOrganizer.OrganizerLogic(
                        Application,
                        domainsFolderName,
                        inboxArchiveFolderName,
                        Globals.Ribbons.Ribbon1.chkChronoSort.Checked
                    );

                    organizerLogic.MessagesRemainingEventHandler += OrganizerLogic_MessagesRemainingEvent;
                }
            }
        }

        private void OrganizerLogic_MessagesRemainingEvent(int messagesRemaining)
        {
            Globals.Ribbons.Ribbon1.btnProcessingQueue.Label = messagesRemaining.ToString();
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
