using Microsoft.Office.Tools.Ribbon;

namespace emailSimplyfier
{
    partial class menueBand : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Erforderliche Designervariable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public menueBand()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Verwendete Ressourcen bereinigen.
        /// </summary>
        /// <param name="disposing">"true", wenn verwaltete Ressourcen gelöscht werden sollen, andernfalls "false".</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Vom Komponenten-Designer generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(menueBand));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.emSimGroup = this.Factory.CreateRibbonGroup();
            this.emSimGallery = this.Factory.CreateRibbonGallery();
            this.archiveEmail = this.Factory.CreateRibbonButton();
            this.processPDF = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.emSimGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.emSimGroup);
            resources.ApplyResources(this.tab1, "tab1");
            this.tab1.Name = "tab1";
            // 
            // emSimGroup
            // 
            this.emSimGroup.Items.Add(this.emSimGallery);
            resources.ApplyResources(this.emSimGroup, "emSimGroup");
            this.emSimGroup.Name = "emSimGroup";
            // 
            // emSimGallery
            // 
            this.emSimGallery.Buttons.Add(this.archiveEmail);
            this.emSimGallery.Buttons.Add(this.processPDF);
            this.emSimGallery.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.emSimGallery.Image = global::emailSimplyfier.Properties.Resources.BIcon;
            resources.ApplyResources(this.emSimGallery, "emSimGallery");
            this.emSimGallery.Name = "emSimGallery";
            this.emSimGallery.ShowImage = true;
            this.emSimGallery.ShowItemSelection = true;
            this.emSimGallery.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.emSimGallery_Click);
            // 
            // archiveEmail
            // 
            this.archiveEmail.Image = global::emailSimplyfier.Properties.Resources.ArchivierenIcon;
            resources.ApplyResources(this.archiveEmail, "archiveEmail");
            this.archiveEmail.Name = "archiveEmail";
            this.archiveEmail.ShowImage = true;
            this.archiveEmail.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ArchiveEmail_Click);
            // 
            // processPDF
            // 
            this.processPDF.Image = global::emailSimplyfier.Properties.Resources.PDFIcon;
            resources.ApplyResources(this.processPDF, "processPDF");
            this.processPDF.Name = "processPDF";
            this.processPDF.ShowImage = true;
            this.processPDF.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ProcessPDF_Click);
            // 
            // menueBand
            // 
            this.Name = "menueBand";
            this.RibbonType = "Microsoft.Outlook.Mail.Read";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.menueBand_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.emSimGroup.ResumeLayout(false);
            this.emSimGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup emSimGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery emSimGallery;
        private RibbonButton archiveEmail;
        private RibbonButton processPDF;
    }

    partial class ThisRibbonCollection
    {
        internal menueBand menueBand
        {
            get { return this.GetRibbon<menueBand>(); }
        }
    }
}
