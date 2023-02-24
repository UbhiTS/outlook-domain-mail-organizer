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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.grpDomainMailOrganizer = this.Factory.CreateRibbonGroup();
            this.enableChronoSort = this.Factory.CreateRibbonCheckBox();
            this.btnOrganize = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.grpDomainMailOrganizer.SuspendLayout();
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
            this.grpDomainMailOrganizer.Items.Add(this.enableChronoSort);
            this.grpDomainMailOrganizer.Items.Add(this.btnOrganize);
            this.grpDomainMailOrganizer.Label = "Domain Mail Organizer";
            this.grpDomainMailOrganizer.Name = "grpDomainMailOrganizer";
            // 
            // enableChronoSort
            // 
            this.enableChronoSort.Checked = true;
            this.enableChronoSort.Label = "Move Newest to Top";
            this.enableChronoSort.Name = "enableChronoSort";
            // 
            // btnOrganize
            // 
            this.btnOrganize.Label = "Organize Now";
            this.btnOrganize.Name = "btnOrganize";
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
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpDomainMailOrganizer;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox enableChronoSort;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnOrganize;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
