using System.Windows.Forms;
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
            Globals.Ribbons.Ribbon1.chkChronoSort.Click += chkChronoSort_Click;
            Globals.Ribbons.Ribbon1.btnOrganizeInbox.Click += btnOrganizeInbox_Click;
            Globals.Ribbons.Ribbon1.btnOrganizeArchive.Click += btnOrganizeArchive_Click;
            
            Application.NewMail += NewMail;
        }

        private void chkChronoSort_Click(object sender, RibbonControlEventArgs e)
        {
            organizerLogic = null;
        }

        private void btnOrganizeInbox_Click(object sender, RibbonControlEventArgs e)
        {
            InitializeOrganizerLogic();

            System.Threading.Thread t = null;

            if (int.Parse(Globals.Ribbons.Ribbon1.ddDays.SelectedItem.Tag.ToString()) == 1) {
                t = new System.Threading.Thread(organizerLogic.ProcessInbox1Day);
            }
            else if (int.Parse(Globals.Ribbons.Ribbon1.ddDays.SelectedItem.Tag.ToString()) == 7)
            {
                t = new System.Threading.Thread(organizerLogic.ProcessInbox7Day);
            }
            else if (int.Parse(Globals.Ribbons.Ribbon1.ddDays.SelectedItem.Tag.ToString()) == 30)
            {
                t = new System.Threading.Thread(organizerLogic.ProcessInbox30Day);
            }
            else
            {
                t = new System.Threading.Thread(organizerLogic.ProcessInboxAll);
            }

            t.SetApartmentState(System.Threading.ApartmentState.STA);
            t.Start();
        }

        private void btnOrganizeArchive_Click(object sender, RibbonControlEventArgs e)
        {
            InitializeOrganizerLogic();

            System.Threading.Thread t = null;

            if (int.Parse(Globals.Ribbons.Ribbon1.ddDays.SelectedItem.Tag.ToString()) == 1)
            {
                t = new System.Threading.Thread(organizerLogic.ProcessArchive1Day);
            }
            else if (int.Parse(Globals.Ribbons.Ribbon1.ddDays.SelectedItem.Tag.ToString()) == 7)
            {
                t = new System.Threading.Thread(organizerLogic.ProcessArchive7Day);
            }
            else if (int.Parse(Globals.Ribbons.Ribbon1.ddDays.SelectedItem.Tag.ToString()) == 30)
            {
                t = new System.Threading.Thread(organizerLogic.ProcessArchive30Day);
            }
            else
            {
                t = new System.Threading.Thread(organizerLogic.ProcessArchiveAll);
            }

            t.SetApartmentState(System.Threading.ApartmentState.STA);
            t.Start();
        }

        private void NewMail()
        {
            InitializeOrganizerLogic();

            System.Threading.Thread t = new System.Threading.Thread(organizerLogic.ProcessInboxUnread);
            t.SetApartmentState(System.Threading.ApartmentState.STA);
            t.Start();
        }

        private void InitializeOrganizerLogic()
        {
            if (organizerLogic == null)
            {
                organizerLogic = new DomainMailOrganizer.OrganizerLogic(
                    Application, 
                    domainsFolderName, 
                    inboxArchiveFolderName, 
                    Globals.Ribbons.Ribbon1.chkChronoSort.Checked
                );

                organizerLogic.StatusUpdate += OrganizerLogic_StatusUpdate;
            }
        }

        private void OrganizerLogic_StatusUpdate(string statusUpdate)
        {
            Globals.Ribbons.Ribbon1.btnProcessingQueue.Label = statusUpdate;
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
