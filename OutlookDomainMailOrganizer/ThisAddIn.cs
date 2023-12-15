using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Linq;

namespace OutlookDomainMailOrganizer
{
    public partial class ThisAddIn
    {
        #region Config Parameters

        string domainsFolderName = "Domains";

        #endregion

        #region Main Methods

        private void ODMOAddIn_Startup(object sender, System.EventArgs e)
        {
            InitPlugin();
        }

        private void ODMOAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #endregion

        #region Constructors and Initializers

        private void InitPlugin()
        {
            Globals.Ribbons.Ribbon1.chkChronoSort.Checked = true;

            Globals.Ribbons.Ribbon1.chkChronoSort.Click += chkChronoSort_Click;
            Globals.Ribbons.Ribbon1.btnOrganizeInbox.Click += btnOrganizeInbox_Click;
            Globals.Ribbons.Ribbon1.btnOrganizeArchive.Click += btnOrganizeArchive_Click;
            Globals.Ribbons.Ribbon1.btnMoveToArchive.Click += btnMoveToArchive_Click;

            Globals.Ribbons.Ribbon1.ddDays.SelectedItem = Globals.Ribbons.Ribbon1.ddDays.Items.Where(x => x.Tag.ToString() == "2").First();

            // Application.NewMail += NewMail;
        }

        private DomainMailOrganizer.OrganizerLogic InitOrganizer()
        {
            DomainMailOrganizer.OrganizerLogic organizerLogic = null;

            if (organizerLogic == null)
            {
                organizerLogic = new DomainMailOrganizer.OrganizerLogic(
                    Application,
                    domainsFolderName,
                    Globals.Ribbons.Ribbon1.chkChronoSort.Checked
                );

                organizerLogic.MessagesRemainingEventHandler += OrganizerLogic_MessagesRemainingEvent;
            }

            return organizerLogic;
        }

        #endregion

        #region Event Handlers

        private void btnOrganizeInbox_Click(object sender, RibbonControlEventArgs e)
        {
            var organizer = InitOrganizer();

            System.Threading.Thread t = null;

            switch (int.Parse(Globals.Ribbons.Ribbon1.ddDays.SelectedItem.Tag.ToString()))
            {
                case 1:
                    t = new System.Threading.Thread(organizer.ProcessInbox1Day);
                    break;
                case 2:
                    t = new System.Threading.Thread(organizer.ProcessInbox2Day);
                    break;
                case 7:
                    t = new System.Threading.Thread(organizer.ProcessInbox7Day);
                    break;
                case 30:
                    t = new System.Threading.Thread(organizer.ProcessInbox30Day);
                    break;
                default:
                    t = new System.Threading.Thread(organizer.ProcessInboxAll);
                    break;
            }

            t.SetApartmentState(System.Threading.ApartmentState.STA);
            t.Start();
        }

        private void btnOrganizeArchive_Click(object sender, RibbonControlEventArgs e)
        {
            var organizer = InitOrganizer();

            System.Threading.Thread t = null;

            switch (int.Parse(Globals.Ribbons.Ribbon1.ddDays.SelectedItem.Tag.ToString()))
            {
                case 1:
                    t = new System.Threading.Thread(organizer.ProcessArchive1Day);
                    break;
                case 2:
                    t = new System.Threading.Thread(organizer.ProcessArchive2Day);
                    break;
                case 7:
                    t = new System.Threading.Thread(organizer.ProcessArchive7Day);
                    break;
                case 30:
                    t = new System.Threading.Thread(organizer.ProcessArchive30Day);
                    break;
                default:
                    t = new System.Threading.Thread(organizer.ProcessArchiveAll);
                    break;
            }

            t.SetApartmentState(System.Threading.ApartmentState.STA);
            t.Start();
        }

        private void btnMoveToArchive_Click(object sender, RibbonControlEventArgs e)
        {
            var organizer = InitOrganizer();
            var t = new System.Threading.Thread(organizer.ArchiveAllInboxItems);
            t.SetApartmentState(System.Threading.ApartmentState.STA);
            t.Start();
        }

        private void chkChronoSort_Click(object sender, RibbonControlEventArgs e)
        {
            var organizer = InitOrganizer();
            var t = new System.Threading.Thread(organizer.ChronoSortFolders);
            t.SetApartmentState(System.Threading.ApartmentState.STA);
            t.Start();
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
