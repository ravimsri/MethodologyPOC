namespace MethodologyPOC
{
    partial class MethodologyRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MethodologyRibbon()
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
            this.tabMethodology = this.Factory.CreateRibbonTab();
            this.grpActions = this.Factory.CreateRibbonGroup();
            this.btnHelloWorld = this.Factory.CreateRibbonButton();
            this.tabMethodology.SuspendLayout();
            this.grpActions.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabMethodology
            // 
            this.tabMethodology.Groups.Add(this.grpActions);
            this.tabMethodology.Label = "Methodology";
            this.tabMethodology.Name = "tabMethodology";
            // 
            // grpActions
            // 
            this.grpActions.Items.Add(this.btnHelloWorld);
            this.grpActions.Label = "Actions";
            this.grpActions.Name = "grpActions";
            // 
            // btnHelloWorld
            // 
            this.btnHelloWorld.Label = "Hello World";
            this.btnHelloWorld.Name = "btnHelloWorld";
            this.btnHelloWorld.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnHelloWorld_Click);
            // 
            // MethodologyRibbon
            // 
            this.Name = "MethodologyRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tabMethodology);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MethodologyRibbon_Load);
            this.tabMethodology.ResumeLayout(false);
            this.tabMethodology.PerformLayout();
            this.grpActions.ResumeLayout(false);
            this.grpActions.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabMethodology;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnHelloWorld;
    }

    partial class ThisRibbonCollection
    {
        internal MethodologyRibbon MethodologyRibbon
        {
            get { return this.GetRibbon<MethodologyRibbon>(); }
        }
    }
}
