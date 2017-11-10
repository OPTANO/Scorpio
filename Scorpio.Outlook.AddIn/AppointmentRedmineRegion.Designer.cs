namespace Scorpio.Outlook.AddIn
{
    using Scorpio.Outlook.AddIn.Cache;
    using Scorpio.Outlook.AddIn.LocalObjects;
    using Scorpio.Outlook.AddIn.Misc;

    [System.ComponentModel.ToolboxItemAttribute(false)]
    partial class AppointmentRedmineRegion : Microsoft.Office.Tools.Outlook.FormRegionBase
    {
        public AppointmentRedmineRegion(Microsoft.Office.Interop.Outlook.FormRegion formRegion)
            : base(Globals.Factory, formRegion)
        {
            this.InitializeComponent();
        }


        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
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
            this.components = new System.ComponentModel.Container();
            this.layoutControl1 = new DevExpress.XtraLayout.LayoutControl();
            this.searchLastIssues = new System.Windows.Forms.TextBox();
            this.lstFavorite = new DevExpress.XtraEditors.ListBoxControl();
            this.favoriteIssuesBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.lstLastUsed = new DevExpress.XtraEditors.ListBoxControl();
            this.lastUsedIssuesBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.issueSelector = new DevExpress.XtraEditors.SearchLookUpEdit();
            this.issueProjectInfoBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.searchLookUpEdit1View = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.colIssueString = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colIssueName = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colProjectName = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colIssueId = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colIssueStatus = new DevExpress.XtraGrid.Columns.GridColumn();
            this.lnkIssue = new System.Windows.Forms.LinkLabel();
            this.lnkProject = new System.Windows.Forms.LinkLabel();
            this.layoutControlGroup1 = new DevExpress.XtraLayout.LayoutControlGroup();
            this.layoutControlGroup2 = new DevExpress.XtraLayout.LayoutControlGroup();
            this.layoutControlItem2 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem7 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlGroup3 = new DevExpress.XtraLayout.LayoutControlGroup();
            this.layoutControlItem1 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem6 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem3 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlGroup4 = new DevExpress.XtraLayout.LayoutControlGroup();
            this.layoutControlItem4 = new DevExpress.XtraLayout.LayoutControlItem();
            this.defaultLookAndFeel1 = new DevExpress.LookAndFeel.DefaultLookAndFeel(this.components);
            this.assignedToMeSource = new System.Windows.Forms.BindingSource(this.components);
            this.layoutControlItem5 = new DevExpress.XtraLayout.LayoutControlItem();
            this.assignedToMeIssuesBinding = new System.Windows.Forms.BindingSource(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).BeginInit();
            this.layoutControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.lstFavorite)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.favoriteIssuesBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lstLastUsed)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lastUsedIssuesBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.issueSelector.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.issueProjectInfoBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.searchLookUpEdit1View)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem7)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem6)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.assignedToMeSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem5)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.assignedToMeIssuesBinding)).BeginInit();
            this.SuspendLayout();
            // 
            // layoutControl1
            // 
            this.layoutControl1.AllowCustomization = false;
            this.layoutControl1.Controls.Add(this.searchLastIssues);
            this.layoutControl1.Controls.Add(this.lstFavorite);
            this.layoutControl1.Controls.Add(this.lstLastUsed);
            this.layoutControl1.Controls.Add(this.issueSelector);
            this.layoutControl1.Controls.Add(this.lnkIssue);
            this.layoutControl1.Controls.Add(this.lnkProject);
            this.layoutControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.layoutControl1.Location = new System.Drawing.Point(0, 0);
            this.layoutControl1.Name = "layoutControl1";
            this.layoutControl1.OptionsCustomizationForm.DesignTimeCustomizationFormPositionAndSize = new System.Drawing.Rectangle(872, 252, 831, 458);
            this.layoutControl1.Root = this.layoutControlGroup1;
            this.layoutControl1.Size = new System.Drawing.Size(662, 177);
            this.layoutControl1.TabIndex = 0;
            this.layoutControl1.Text = "layoutControl1";
            // 
            // searchLastIssues
            // 
            this.searchLastIssues.Location = new System.Drawing.Point(336, 43);
            this.searchLastIssues.Name = "searchLastIssues";
            this.searchLastIssues.Size = new System.Drawing.Size(93, 20);
            this.searchLastIssues.TabIndex = 14;
            this.searchLastIssues.TextChanged += new System.EventHandler(this.searchLastIssues_TextChanged);
            // 
            // lstFavorite
            // 
            this.lstFavorite.DataSource = this.favoriteIssuesBindingSource;
            this.lstFavorite.DisplayMember = "DisplayValue";
            this.lstFavorite.Location = new System.Drawing.Point(457, 43);
            this.lstFavorite.Name = "lstFavorite";
            this.lstFavorite.Size = new System.Drawing.Size(181, 110);
            this.lstFavorite.SortOrder = System.Windows.Forms.SortOrder.Ascending;
            this.lstFavorite.StyleController = this.layoutControl1;
            this.lstFavorite.TabIndex = 13;
            this.lstFavorite.ValueMember = "Id";
            this.lstFavorite.SelectedIndexChanged += new System.EventHandler(this.LstFavorite_SelectedValueChanged);
            // 
            // favoriteIssuesBindingSource
            // 
            this.favoriteIssuesBindingSource.AllowNew = false;
            // 
            // lstLastUsed
            // 
            this.lstLastUsed.DataSource = this.lastUsedIssuesBindingSource;
            this.lstLastUsed.DisplayMember = "DisplayValue";
            this.lstLastUsed.Location = new System.Drawing.Point(253, 67);
            this.lstLastUsed.Name = "lstLastUsed";
            this.lstLastUsed.Size = new System.Drawing.Size(176, 86);
            this.lstLastUsed.StyleController = this.layoutControl1;
            this.lstLastUsed.TabIndex = 12;
            this.lstLastUsed.ValueMember = "Id";
            this.lstLastUsed.SelectedValueChanged += new System.EventHandler(this.LstLastUsed_SelectedValueChanged);
            // 
            // lastUsedIssuesBindingSource
            // 
            this.lastUsedIssuesBindingSource.AllowNew = false;
            // 
            // issueSelector
            // 
            this.issueSelector.Location = new System.Drawing.Point(107, 43);
            this.issueSelector.Name = "issueSelector";
            this.issueSelector.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.issueSelector.Properties.DataSource = this.issueProjectInfoBindingSource;
            this.issueSelector.Properties.DisplayMember = "DisplayValue";
            this.issueSelector.Properties.ValueMember = "Id";
            this.issueSelector.Properties.View = this.searchLookUpEdit1View;
            this.issueSelector.Size = new System.Drawing.Size(118, 20);
            this.issueSelector.StyleController = this.layoutControl1;
            this.issueSelector.TabIndex = 11;
            this.issueSelector.Popup += new System.EventHandler(this.issueSelector_Popup);
            this.issueSelector.EditValueChanged += new System.EventHandler(this.IssueSelector_EditValueChanged);
            // 
            // issueProjectInfoBindingSource
            // 
            this.issueProjectInfoBindingSource.AllowNew = false;
            // 
            // searchLookUpEdit1View
            // 
            this.searchLookUpEdit1View.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.colIssueString,
            this.colIssueName,
            this.colProjectName,
            this.colIssueId,
            this.colIssueStatus});
            this.searchLookUpEdit1View.FocusRectStyle = DevExpress.XtraGrid.Views.Grid.DrawFocusRectStyle.RowFocus;
            this.searchLookUpEdit1View.GroupCount = 1;
            this.searchLookUpEdit1View.Name = "searchLookUpEdit1View";
            this.searchLookUpEdit1View.OptionsBehavior.AllowFixedGroups = DevExpress.Utils.DefaultBoolean.True;
            this.searchLookUpEdit1View.OptionsBehavior.AutoExpandAllGroups = true;
            this.searchLookUpEdit1View.OptionsSelection.EnableAppearanceFocusedCell = false;
            this.searchLookUpEdit1View.OptionsView.ShowAutoFilterRow = true;
            this.searchLookUpEdit1View.OptionsView.ShowGroupPanel = false;
            this.searchLookUpEdit1View.SortInfo.AddRange(new DevExpress.XtraGrid.Columns.GridColumnSortInfo[] {
            new DevExpress.XtraGrid.Columns.GridColumnSortInfo(this.colProjectName, DevExpress.Data.ColumnSortOrder.Ascending),
            new DevExpress.XtraGrid.Columns.GridColumnSortInfo(this.colIssueStatus, DevExpress.Data.ColumnSortOrder.Descending)});
            // 
            // colIssueString
            // 
            this.colIssueString.Caption = "Id";
            this.colIssueString.FieldName = "IssueString";
            this.colIssueString.Name = "colIssueString";
            this.colIssueString.Visible = true;
            this.colIssueString.VisibleIndex = 1;
            // 
            // colIssueName
            // 
            this.colIssueName.Caption = "Issue";
            this.colIssueName.FieldName = "Name";
            this.colIssueName.Name = "colIssueName";
            this.colIssueName.Visible = true;
            this.colIssueName.VisibleIndex = 0;
            this.colIssueName.Width = 339;
            // 
            // colProjectName
            // 
            this.colProjectName.Caption = "Projekt";
            this.colProjectName.FieldName = "ProjectShortName";
            this.colProjectName.Name = "colProjectName";
            this.colProjectName.OptionsColumn.AllowGroup = DevExpress.Utils.DefaultBoolean.True;
            this.colProjectName.Visible = true;
            this.colProjectName.VisibleIndex = 2;
            // 
            // colIssueId
            // 
            this.colIssueId.FieldName = "IssueId";
            this.colIssueId.Name = "colIssueId";
            // 
            // colIssueStatus
            // 
            this.colIssueStatus.Caption = "Status";
            this.colIssueStatus.FieldName = "StatusString";
            this.colIssueStatus.FilterMode = DevExpress.XtraGrid.ColumnFilterMode.DisplayText;
            this.colIssueStatus.Name = "colIssueStatus";
            this.colIssueStatus.Visible = true;
            this.colIssueStatus.VisibleIndex = 2;
            // 
            // lnkIssue
            // 
            this.lnkIssue.Location = new System.Drawing.Point(107, 91);
            this.lnkIssue.Name = "lnkIssue";
            this.lnkIssue.Size = new System.Drawing.Size(118, 20);
            this.lnkIssue.TabIndex = 1;
            this.lnkIssue.TabStop = true;
            this.lnkIssue.Text = "linkLabel2";
            this.lnkIssue.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.LnkProject_LinkClicked);
            // 
            // lnkProject
            // 
            this.lnkProject.Location = new System.Drawing.Point(107, 67);
            this.lnkProject.Name = "lnkProject";
            this.lnkProject.Size = new System.Drawing.Size(118, 20);
            this.lnkProject.TabIndex = 9;
            this.lnkProject.TabStop = true;
            this.lnkProject.Text = "linkLabel1";
            this.lnkProject.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.LnkProject_LinkClicked);
            // 
            // layoutControlGroup1
            // 
            this.layoutControlGroup1.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            this.layoutControlGroup1.GroupBordersVisible = false;
            this.layoutControlGroup1.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.layoutControlGroup2,
            this.layoutControlGroup3,
            this.layoutControlGroup4});
            this.layoutControlGroup1.Location = new System.Drawing.Point(0, 0);
            this.layoutControlGroup1.Name = "Root";
            this.layoutControlGroup1.Size = new System.Drawing.Size(662, 177);
            this.layoutControlGroup1.Text = "Issue auswählen";
            // 
            // layoutControlGroup2
            // 
            this.layoutControlGroup2.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.layoutControlItem2,
            this.layoutControlItem7});
            this.layoutControlGroup2.Location = new System.Drawing.Point(229, 0);
            this.layoutControlGroup2.Name = "layoutControlGroup2";
            this.layoutControlGroup2.Size = new System.Drawing.Size(204, 157);
            this.layoutControlGroup2.Text = "Zuletzt benutzt";
            // 
            // layoutControlItem2
            // 
            this.layoutControlItem2.Control = this.lstLastUsed;
            this.layoutControlItem2.Location = new System.Drawing.Point(0, 24);
            this.layoutControlItem2.Name = "layoutControlItem2";
            this.layoutControlItem2.Size = new System.Drawing.Size(180, 90);
            this.layoutControlItem2.TextLocation = DevExpress.Utils.Locations.Top;
            this.layoutControlItem2.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem2.TextVisible = false;
            // 
            // layoutControlItem7
            // 
            this.layoutControlItem7.Control = this.searchLastIssues;
            this.layoutControlItem7.Location = new System.Drawing.Point(0, 0);
            this.layoutControlItem7.Name = "layoutControlItem7";
            this.layoutControlItem7.Size = new System.Drawing.Size(180, 24);
            this.layoutControlItem7.Text = "Suchen:";
            this.layoutControlItem7.TextSize = new System.Drawing.Size(80, 13);
            // 
            // layoutControlGroup3
            // 
            this.layoutControlGroup3.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.layoutControlItem1,
            this.layoutControlItem6,
            this.layoutControlItem3});
            this.layoutControlGroup3.Location = new System.Drawing.Point(0, 0);
            this.layoutControlGroup3.Name = "layoutControlGroup3";
            this.layoutControlGroup3.Size = new System.Drawing.Size(229, 157);
            this.layoutControlGroup3.Text = "Issue + Projekt";
            // 
            // layoutControlItem1
            // 
            this.layoutControlItem1.Control = this.lnkIssue;
            this.layoutControlItem1.Location = new System.Drawing.Point(0, 48);
            this.layoutControlItem1.Name = "layoutControlItem1";
            this.layoutControlItem1.OptionsTableLayoutItem.ColumnIndex = 1;
            this.layoutControlItem1.Size = new System.Drawing.Size(205, 66);
            this.layoutControlItem1.Text = "Issue";
            this.layoutControlItem1.TextSize = new System.Drawing.Size(80, 13);
            // 
            // layoutControlItem6
            // 
            this.layoutControlItem6.Control = this.lnkProject;
            this.layoutControlItem6.Location = new System.Drawing.Point(0, 24);
            this.layoutControlItem6.Name = "layoutControlItem6";
            this.layoutControlItem6.Size = new System.Drawing.Size(205, 24);
            this.layoutControlItem6.Text = "Projekt";
            this.layoutControlItem6.TextSize = new System.Drawing.Size(80, 13);
            // 
            // layoutControlItem3
            // 
            this.layoutControlItem3.Control = this.issueSelector;
            this.layoutControlItem3.Location = new System.Drawing.Point(0, 0);
            this.layoutControlItem3.Name = "layoutControlItem3";
            this.layoutControlItem3.OptionsTableLayoutItem.RowIndex = 1;
            this.layoutControlItem3.Size = new System.Drawing.Size(205, 24);
            this.layoutControlItem3.Text = "Issue auswählen";
            this.layoutControlItem3.TextSize = new System.Drawing.Size(80, 13);
            // 
            // layoutControlGroup4
            // 
            this.layoutControlGroup4.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.layoutControlItem4});
            this.layoutControlGroup4.Location = new System.Drawing.Point(433, 0);
            this.layoutControlGroup4.Name = "layoutControlGroup4";
            this.layoutControlGroup4.Size = new System.Drawing.Size(209, 157);
            this.layoutControlGroup4.Text = "Favoriten";
            // 
            // layoutControlItem4
            // 
            this.layoutControlItem4.Control = this.lstFavorite;
            this.layoutControlItem4.Location = new System.Drawing.Point(0, 0);
            this.layoutControlItem4.Name = "layoutControlItem4";
            this.layoutControlItem4.Size = new System.Drawing.Size(185, 114);
            this.layoutControlItem4.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem4.TextVisible = false;
            // 
            // defaultLookAndFeel1
            // 
            this.defaultLookAndFeel1.LookAndFeel.SkinName = "Office 2013 Dark Gray";
            // 
            // layoutControlItem5
            // 
            this.layoutControlItem5.Location = new System.Drawing.Point(0, 0);
            this.layoutControlItem5.Name = "layoutControlItem5";
            this.layoutControlItem5.Size = new System.Drawing.Size(0, 0);
            this.layoutControlItem5.TextSize = new System.Drawing.Size(50, 20);
            // 
            // AppointmentRedmineRegion
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.layoutControl1);
            this.Name = "AppointmentRedmineRegion";
            this.Size = new System.Drawing.Size(662, 177);
            this.FormRegionShowing += new System.EventHandler(this.AppointmentRedmineRegion_FormRegionShowing);
            this.FormRegionClosed += new System.EventHandler(this.AppointmentRedmineRegion_FormRegionClosed);
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).EndInit();
            this.layoutControl1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.lstFavorite)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.favoriteIssuesBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lstLastUsed)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lastUsedIssuesBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.issueSelector.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.issueProjectInfoBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.searchLookUpEdit1View)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem7)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem6)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.assignedToMeSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem5)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.assignedToMeIssuesBinding)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        #region Form Region Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private static void InitializeManifest(Microsoft.Office.Tools.Outlook.FormRegionManifest manifest, Microsoft.Office.Tools.Outlook.Factory factory)
        {
            manifest.FormRegionName = "ORCONOMY Tool - Redmine";
            manifest.FormRegionType = Microsoft.Office.Tools.Outlook.FormRegionType.Adjoining;
            manifest.Title = "ORCONOMY Tool - Redmine";

        }

        #endregion

        private DevExpress.XtraLayout.LayoutControl layoutControl1;
        private DevExpress.XtraLayout.LayoutControlGroup layoutControlGroup1;
        private System.Windows.Forms.LinkLabel lnkIssue;
        private System.Windows.Forms.LinkLabel lnkProject;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem6;
        private System.Windows.Forms.BindingSource issueProjectInfoBindingSource;
        private DevExpress.XtraEditors.SearchLookUpEdit issueSelector;
        private DevExpress.XtraGrid.Views.Grid.GridView searchLookUpEdit1View;
        private DevExpress.XtraGrid.Columns.GridColumn colIssueName;
        private DevExpress.XtraGrid.Columns.GridColumn colProjectName;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem3;
        private DevExpress.XtraGrid.Columns.GridColumn colIssueString;
        private DevExpress.XtraGrid.Columns.GridColumn colIssueId;
        private DevExpress.LookAndFeel.DefaultLookAndFeel defaultLookAndFeel1;
        private DevExpress.XtraEditors.ListBoxControl lstLastUsed;
        private DevExpress.XtraLayout.LayoutControlGroup layoutControlGroup2;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem2;
        private DevExpress.XtraLayout.LayoutControlGroup layoutControlGroup3;
        private System.Windows.Forms.BindingSource lastUsedIssuesBindingSource;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem1;
        private DevExpress.XtraEditors.ListBoxControl lstFavorite;
        private DevExpress.XtraLayout.LayoutControlGroup layoutControlGroup4;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem4;
        private System.Windows.Forms.BindingSource favoriteIssuesBindingSource;
        private System.Windows.Forms.BindingSource assignedToMeSource;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem5;
        private System.Windows.Forms.BindingSource assignedToMeIssuesBinding;
        private System.Windows.Forms.TextBox searchLastIssues;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem7;
        private DevExpress.XtraGrid.Columns.GridColumn colIssueStatus;

        public partial class AppointmentRedmineRegionFactory : Microsoft.Office.Tools.Outlook.IFormRegionFactory
        {
            public event Microsoft.Office.Tools.Outlook.FormRegionInitializingEventHandler FormRegionInitializing;

            private Microsoft.Office.Tools.Outlook.FormRegionManifest _Manifest;

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            public AppointmentRedmineRegionFactory()
            {
                this._Manifest = Globals.Factory.CreateFormRegionManifest();
                AppointmentRedmineRegion.InitializeManifest(this._Manifest, Globals.Factory);
                this.FormRegionInitializing += new Microsoft.Office.Tools.Outlook.FormRegionInitializingEventHandler(this.AppointmentRedmineRegionFactory_FormRegionInitializing);
            }

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            public Microsoft.Office.Tools.Outlook.FormRegionManifest Manifest
            {
                get
                {
                    return this._Manifest;
                }
            }

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            Microsoft.Office.Tools.Outlook.IFormRegion Microsoft.Office.Tools.Outlook.IFormRegionFactory.CreateFormRegion(Microsoft.Office.Interop.Outlook.FormRegion formRegion)
            {
                AppointmentRedmineRegion form = new AppointmentRedmineRegion(formRegion);
                form.Factory = this;
                return form;
            }

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            byte[] Microsoft.Office.Tools.Outlook.IFormRegionFactory.GetFormRegionStorage(object outlookItem, Microsoft.Office.Interop.Outlook.OlFormRegionMode formRegionMode, Microsoft.Office.Interop.Outlook.OlFormRegionSize formRegionSize)
            {
                throw new System.NotSupportedException();
            }

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            bool Microsoft.Office.Tools.Outlook.IFormRegionFactory.IsDisplayedForItem(object outlookItem, Microsoft.Office.Interop.Outlook.OlFormRegionMode formRegionMode, Microsoft.Office.Interop.Outlook.OlFormRegionSize formRegionSize)
            {
                if (this.FormRegionInitializing != null)
                {
                    Microsoft.Office.Tools.Outlook.FormRegionInitializingEventArgs cancelArgs = Globals.Factory.CreateFormRegionInitializingEventArgs(outlookItem, formRegionMode, formRegionSize, false);
                    this.FormRegionInitializing(this, cancelArgs);
                    return !cancelArgs.Cancel;
                }
                else
                {
                    return true;
                }
            }

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            Microsoft.Office.Tools.Outlook.FormRegionKindConstants Microsoft.Office.Tools.Outlook.IFormRegionFactory.Kind
            {
                get
                {
                    return Microsoft.Office.Tools.Outlook.FormRegionKindConstants.WindowsForms;
                }
            }
        }
    }

    partial class WindowFormRegionCollection
    {
        internal AppointmentRedmineRegion AppointmentRedmineRegion
        {
            get
            {
                foreach (var item in this)
                {
                    if (item.GetType() == typeof(AppointmentRedmineRegion))
                        return (AppointmentRedmineRegion)item;
                }
                return null;
            }
        }
    }
}
