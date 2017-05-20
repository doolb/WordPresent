namespace WordPresent
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
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnDataBase = this.Factory.CreateRibbonButton();
            this.cmbTables = this.Factory.CreateRibbonComboBox();
            this.btnPresent = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnDataBase);
            this.group1.Items.Add(this.cmbTables);
            this.group1.Items.Add(this.btnPresent);
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // btnDataBase
            // 
            this.btnDataBase.Label = "Select Database File ";
            this.btnDataBase.Name = "btnDataBase";
            this.btnDataBase.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Database_Click);
            // 
            // cmbTables
            // 
            this.cmbTables.Label = "Table";
            this.cmbTables.Name = "cmbTables";
            this.cmbTables.Text = null;
            this.cmbTables.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cmbTables_TextChanged);

            // 
            // btnPresent
            // 
            this.btnPresent.Label = "Present";
            this.btnPresent.Name = "btnPresent";
            this.btnPresent.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Present_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDataBase;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox cmbTables;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPresent;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
