namespace MailClassificationAddin
{
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon()
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
            this.tab = this.Factory.CreateRibbonTab();
            this.groupClassify = this.Factory.CreateRibbonGroup();
            this.buttonClassify = this.Factory.CreateRibbonButton();
            this.buttonTrain = this.Factory.CreateRibbonButton();
            this.tab.SuspendLayout();
            this.groupClassify.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab
            // 
            this.tab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab.Groups.Add(this.groupClassify);
            this.tab.Label = "TabAddIns";
            this.tab.Name = "tab";
            // 
            // groupClassify
            // 
            this.groupClassify.Items.Add(this.buttonClassify);
            this.groupClassify.Items.Add(this.buttonTrain);
            this.groupClassify.Label = "Mail Classification";
            this.groupClassify.Name = "groupClassify";
            // 
            // buttonClassify
            // 
            this.buttonClassify.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonClassify.Image = global::MailClassificationAddin.Resources.sort;
            this.buttonClassify.Label = "Classify";
            this.buttonClassify.Name = "buttonClassify";
            this.buttonClassify.ShowImage = true;
            this.buttonClassify.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonClassify_Click);
            // 
            // buttonTrain
            // 
            this.buttonTrain.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonTrain.Image = global::MailClassificationAddin.Resources.folder;
            this.buttonTrain.Label = "Train";
            this.buttonTrain.Name = "buttonTrain";
            this.buttonTrain.ShowImage = true;
            this.buttonTrain.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonTrain_Click);
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Outlook.Explorer, Microsoft.Outlook.Mail.Read";
            this.Tabs.Add(this.tab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.tab.ResumeLayout(false);
            this.tab.PerformLayout();
            this.groupClassify.ResumeLayout(false);
            this.groupClassify.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupClassify;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonClassify;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonTrain;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon1
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
