namespace OutlookDomainMailOrganizer
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon1));
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl3 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl4 = this.Factory.CreateRibbonDropDownItem();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.grpDomainMailOrganizer = this.Factory.CreateRibbonGroup();
            this.btnRefresh = this.Factory.CreateRibbonButton();
            this.separator4 = this.Factory.CreateRibbonSeparator();
            this.chkChronoSort = this.Factory.CreateRibbonToggleButton();
            this.separator3 = this.Factory.CreateRibbonSeparator();
            this.box2 = this.Factory.CreateRibbonBox();
            this.ddDays = this.Factory.CreateRibbonDropDown();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.btnOrganizeInbox = this.Factory.CreateRibbonButton();
            this.btnOrganizeArchive = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.btnProcessingQueue = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.grpDomainMailOrganizer.SuspendLayout();
            this.box2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.grpDomainMailOrganizer);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // grpDomainMailOrganizer
            // 
            this.grpDomainMailOrganizer.Items.Add(this.btnRefresh);
            this.grpDomainMailOrganizer.Items.Add(this.separator4);
            this.grpDomainMailOrganizer.Items.Add(this.chkChronoSort);
            this.grpDomainMailOrganizer.Items.Add(this.separator3);
            this.grpDomainMailOrganizer.Items.Add(this.box2);
            this.grpDomainMailOrganizer.Items.Add(this.separator2);
            this.grpDomainMailOrganizer.Items.Add(this.btnOrganizeInbox);
            this.grpDomainMailOrganizer.Items.Add(this.btnOrganizeArchive);
            this.grpDomainMailOrganizer.Items.Add(this.separator1);
            this.grpDomainMailOrganizer.Items.Add(this.btnProcessingQueue);
            this.grpDomainMailOrganizer.Label = "Domain Mail Organizer";
            this.grpDomainMailOrganizer.Name = "grpDomainMailOrganizer";
            // 
            // btnRefresh
            // 
            this.btnRefresh.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnRefresh.Image = ((System.Drawing.Image)(resources.GetObject("btnRefresh.Image")));
            this.btnRefresh.Label = "Reload Config";
            this.btnRefresh.Name = "btnRefresh";
            this.btnRefresh.ShowImage = true;
            // 
            // separator4
            // 
            this.separator4.Name = "separator4";
            // 
            // chkChronoSort
            // 
            this.chkChronoSort.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.chkChronoSort.Image = ((System.Drawing.Image)(resources.GetObject("chkChronoSort.Image")));
            this.chkChronoSort.Label = "Newest to Top";
            this.chkChronoSort.Name = "chkChronoSort";
            this.chkChronoSort.ShowImage = true;
            // 
            // separator3
            // 
            this.separator3.Name = "separator3";
            // 
            // box2
            // 
            this.box2.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            this.box2.Items.Add(this.ddDays);
            this.box2.Name = "box2";
            // 
            // ddDays
            // 
            this.ddDays.Image = ((System.Drawing.Image)(resources.GetObject("ddDays.Image")));
            ribbonDropDownItemImpl1.Label = "1 Day";
            ribbonDropDownItemImpl1.Tag = "1";
            ribbonDropDownItemImpl2.Label = "7 Days";
            ribbonDropDownItemImpl2.Tag = "7";
            ribbonDropDownItemImpl3.Label = "30 Days";
            ribbonDropDownItemImpl3.Tag = "30";
            ribbonDropDownItemImpl4.Label = "All";
            ribbonDropDownItemImpl4.Tag = "0";
            this.ddDays.Items.Add(ribbonDropDownItemImpl1);
            this.ddDays.Items.Add(ribbonDropDownItemImpl2);
            this.ddDays.Items.Add(ribbonDropDownItemImpl3);
            this.ddDays.Items.Add(ribbonDropDownItemImpl4);
            this.ddDays.Label = "Days";
            this.ddDays.Name = "ddDays";
            this.ddDays.ShowImage = true;
            this.ddDays.ShowLabel = false;
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // btnOrganizeInbox
            // 
            this.btnOrganizeInbox.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnOrganizeInbox.Image = ((System.Drawing.Image)(resources.GetObject("btnOrganizeInbox.Image")));
            this.btnOrganizeInbox.Label = "Process Inbox";
            this.btnOrganizeInbox.Name = "btnOrganizeInbox";
            this.btnOrganizeInbox.ShowImage = true;
            this.btnOrganizeInbox.SuperTip = "(24 hours)";
            // 
            // btnOrganizeArchive
            // 
            this.btnOrganizeArchive.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnOrganizeArchive.Image = ((System.Drawing.Image)(resources.GetObject("btnOrganizeArchive.Image")));
            this.btnOrganizeArchive.Label = "Process Archive";
            this.btnOrganizeArchive.Name = "btnOrganizeArchive";
            this.btnOrganizeArchive.ShowImage = true;
            this.btnOrganizeArchive.SuperTip = "(30 days)";
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // btnProcessingQueue
            // 
            this.btnProcessingQueue.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnProcessingQueue.Image = ((System.Drawing.Image)(resources.GetObject("btnProcessingQueue.Image")));
            this.btnProcessingQueue.Label = "0";
            this.btnProcessingQueue.Name = "btnProcessingQueue";
            this.btnProcessingQueue.ShowImage = true;
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.grpDomainMailOrganizer.ResumeLayout(false);
            this.grpDomainMailOrganizer.PerformLayout();
            this.box2.ResumeLayout(false);
            this.box2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpDomainMailOrganizer;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnOrganizeInbox;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton chkChronoSort;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnOrganizeArchive;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ddDays;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator3;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box2;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnProcessingQueue;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRefresh;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator4;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
