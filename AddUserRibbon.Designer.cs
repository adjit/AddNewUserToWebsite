namespace SendWebUsername
{
    partial class AddUserRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public AddUserRibbon()
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
            this.addUserTab = this.Factory.CreateRibbonTab();
            this.addUserGroup = this.Factory.CreateRibbonGroup();
            this.addUserButton = this.Factory.CreateRibbonButton();
            this.addUserTab.SuspendLayout();
            this.addUserGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // addUserTab
            // 
            this.addUserTab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.addUserTab.Groups.Add(this.addUserGroup);
            this.addUserTab.Label = "Add New User";
            this.addUserTab.Name = "addUserTab";
            // 
            // addUserGroup
            // 
            this.addUserGroup.Items.Add(this.addUserButton);
            this.addUserGroup.Label = "Add User Group";
            this.addUserGroup.Name = "addUserGroup";
            // 
            // addUserButton
            // 
            this.addUserButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.addUserButton.Image = global::SendWebUsername.Properties.Resources.add_user;
            this.addUserButton.Label = "Add User";
            this.addUserButton.Name = "addUserButton";
            this.addUserButton.ScreenTip = "Add New Website User";
            this.addUserButton.ShowImage = true;
            this.addUserButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.addUserButton_Click);
            // 
            // AddUserRibbon
            // 
            this.Name = "AddUserRibbon";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.addUserTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.AddUserRibbon_Load);
            this.addUserTab.ResumeLayout(false);
            this.addUserTab.PerformLayout();
            this.addUserGroup.ResumeLayout(false);
            this.addUserGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab addUserTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup addUserGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton addUserButton;
    }

    partial class ThisRibbonCollection
    {
        internal AddUserRibbon AddUserRibbon
        {
            get { return this.GetRibbon<AddUserRibbon>(); }
        }
    }
}
