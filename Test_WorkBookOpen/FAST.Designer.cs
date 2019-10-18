namespace Test_WorkBookOpen
{
    partial class FAST : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public FAST()
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
            this.grpInformation = this.Factory.CreateRibbonGroup();
            this.label1 = this.Factory.CreateRibbonLabel();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.label2 = this.Factory.CreateRibbonLabel();
            this.grpPromoFilter = this.Factory.CreateRibbonGroup();
            this.ddnCountry = this.Factory.CreateRibbonDropDown();
            this.label7 = this.Factory.CreateRibbonLabel();
            this.ddnDeviceType = this.Factory.CreateRibbonDropDown();
            this.grpFilter = this.Factory.CreateRibbonGroup();
            this.ddnScenario = this.Factory.CreateRibbonDropDown();
            this.label3 = this.Factory.CreateRibbonLabel();
            this.ddnInputType = this.Factory.CreateRibbonDropDown();
            this.separator5 = this.Factory.CreateRibbonSeparator();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.ddnProductLine = this.Factory.CreateRibbonDropDown();
            this.label9 = this.Factory.CreateRibbonLabel();
            this.ddnInterval = this.Factory.CreateRibbonDropDown();
            this.separator6 = this.Factory.CreateRibbonSeparator();
            this.ddnCurrency = this.Factory.CreateRibbonDropDown();
            this.grpOperations = this.Factory.CreateRibbonGroup();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.separator7 = this.Factory.CreateRibbonSeparator();
            this.separator8 = this.Factory.CreateRibbonSeparator();
            this.separator3 = this.Factory.CreateRibbonSeparator();
            this.separator11 = this.Factory.CreateRibbonSeparator();
            this.separator10 = this.Factory.CreateRibbonSeparator();
            this.separator9 = this.Factory.CreateRibbonSeparator();
            this.separator4 = this.Factory.CreateRibbonSeparator();
            this.grpReports = this.Factory.CreateRibbonGroup();
            this.label5 = this.Factory.CreateRibbonLabel();
            this.grpDebug = this.Factory.CreateRibbonGroup();
            this.lblUrl = this.Factory.CreateRibbonEditBox();
            this.lblAliasID = this.Factory.CreateRibbonEditBox();
            this.grpVersion = this.Factory.CreateRibbonGroup();
            this.lblVersion = this.Factory.CreateRibbonLabel();
            this.label6 = this.Factory.CreateRibbonLabel();
            this.button1 = this.Factory.CreateRibbonButton();
            this.btnMenuInitialize = this.Factory.CreateRibbonMenu();
            this.btnDownloadData = this.Factory.CreateRibbonButton();
            this.btnUploadData = this.Factory.CreateRibbonButton();
            this.btnRefresh = this.Factory.CreateRibbonButton();
            this.btnRefreshBransonData = this.Factory.CreateRibbonButton();
            this.btnContactSupport = this.Factory.CreateRibbonButton();
            this.mnuReports = this.Factory.CreateRibbonMenu();
            this.menuVarianceReport = this.Factory.CreateRibbonMenu();
            this.menuAuditReport = this.Factory.CreateRibbonMenu();
            this.menuStatistics = this.Factory.CreateRibbonMenu();
            this.tab1.SuspendLayout();
            this.grpInformation.SuspendLayout();
            this.group1.SuspendLayout();
            this.grpPromoFilter.SuspendLayout();
            this.grpFilter.SuspendLayout();
            this.grpOperations.SuspendLayout();
            this.grpReports.SuspendLayout();
            this.grpDebug.SuspendLayout();
            this.grpVersion.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.grpInformation);
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.grpPromoFilter);
            this.tab1.Groups.Add(this.grpFilter);
            this.tab1.Groups.Add(this.grpOperations);
            this.tab1.Groups.Add(this.grpReports);
            this.tab1.Groups.Add(this.grpDebug);
            this.tab1.Groups.Add(this.grpVersion);
            this.tab1.Label = "FAST";
            this.tab1.Name = "tab1";
            // 
            // grpInformation
            // 
            this.grpInformation.Items.Add(this.label1);
            this.grpInformation.Items.Add(this.button1);
            this.grpInformation.Label = "Information";
            this.grpInformation.Name = "grpInformation";
            // 
            // label1
            // 
            this.label1.Label = "   ";
            this.label1.Name = "label1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.label2);
            this.group1.Items.Add(this.btnMenuInitialize);
            this.group1.Label = "Initialize";
            this.group1.Name = "group1";
            // 
            // label2
            // 
            this.label2.Label = "   ";
            this.label2.Name = "label2";
            // 
            // grpPromoFilter
            // 
            this.grpPromoFilter.Items.Add(this.ddnCountry);
            this.grpPromoFilter.Items.Add(this.label7);
            this.grpPromoFilter.Items.Add(this.ddnDeviceType);
            this.grpPromoFilter.Label = "Filter Selection";
            this.grpPromoFilter.Name = "grpPromoFilter";
            this.grpPromoFilter.Visible = false;
            // 
            // ddnCountry
            // 
            this.ddnCountry.Label = "Country  ";
            this.ddnCountry.Name = "ddnCountry";
            this.ddnCountry.SizeString = "Amazon Devices Tool";
            // 
            // label7
            // 
            this.label7.Label = "   ";
            this.label7.Name = "label7";
            // 
            // ddnDeviceType
            // 
            this.ddnDeviceType.Label = "Device Type";
            this.ddnDeviceType.Name = "ddnDeviceType";
            this.ddnDeviceType.SizeString = "Amazon Devices Tool";
            // 
            // grpFilter
            // 
            this.grpFilter.Items.Add(this.ddnScenario);
            this.grpFilter.Items.Add(this.label3);
            this.grpFilter.Items.Add(this.ddnInputType);
            this.grpFilter.Items.Add(this.separator5);
            this.grpFilter.Items.Add(this.separator1);
            this.grpFilter.Items.Add(this.ddnProductLine);
            this.grpFilter.Items.Add(this.label9);
            this.grpFilter.Items.Add(this.ddnInterval);
            this.grpFilter.Items.Add(this.separator6);
            this.grpFilter.Items.Add(this.ddnCurrency);
            this.grpFilter.Label = "Filter Selection";
            this.grpFilter.Name = "grpFilter";
            // 
            // ddnScenario
            // 
            this.ddnScenario.Label = "Scenario    ";
            this.ddnScenario.Name = "ddnScenario";
            this.ddnScenario.SizeString = "Amazon Devices Input";
            this.ddnScenario.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ddnScenario_SelectionChanged);
            // 
            // label3
            // 
            this.label3.Label = "   ";
            this.label3.Name = "label3";
            // 
            // ddnInputType
            // 
            this.ddnInputType.Label = "Input Type";
            this.ddnInputType.Name = "ddnInputType";
            this.ddnInputType.SizeString = "Amazon Devices Input";
            this.ddnInputType.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ddnInputType_SelectionChanged);
            // 
            // separator5
            // 
            this.separator5.Name = "separator5";
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // ddnProductLine
            // 
            this.ddnProductLine.Label = "Product Line  ";
            this.ddnProductLine.Name = "ddnProductLine";
            this.ddnProductLine.SizeString = "Amazon Devices Input";
            this.ddnProductLine.Visible = false;
            this.ddnProductLine.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ddnProductLine_SelectionChanged);
            // 
            // label9
            // 
            this.label9.Label = "   ";
            this.label9.Name = "label9";
            // 
            // ddnInterval
            // 
            this.ddnInterval.Enabled = false;
            this.ddnInterval.Label = "Interval";
            this.ddnInterval.Name = "ddnInterval";
            this.ddnInterval.SizeString = "Amazon Devices Input";
            this.ddnInterval.Visible = false;
            this.ddnInterval.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ddnInterval_SelectionChanged);
            // 
            // separator6
            // 
            this.separator6.Name = "separator6";
            // 
            // ddnCurrency
            // 
            this.ddnCurrency.Label = "Currency";
            this.ddnCurrency.Name = "ddnCurrency";
            this.ddnCurrency.SizeString = "Amazon Devices Input";
            // 
            // grpOperations
            // 
            this.grpOperations.Items.Add(this.btnDownloadData);
            this.grpOperations.Items.Add(this.separator2);
            this.grpOperations.Items.Add(this.separator7);
            this.grpOperations.Items.Add(this.btnUploadData);
            this.grpOperations.Items.Add(this.separator8);
            this.grpOperations.Items.Add(this.separator3);
            this.grpOperations.Items.Add(this.btnRefresh);
            this.grpOperations.Items.Add(this.separator11);
            this.grpOperations.Items.Add(this.separator10);
            this.grpOperations.Items.Add(this.btnRefreshBransonData);
            this.grpOperations.Items.Add(this.separator9);
            this.grpOperations.Items.Add(this.separator4);
            this.grpOperations.Items.Add(this.btnContactSupport);
            this.grpOperations.Label = "Operations";
            this.grpOperations.Name = "grpOperations";
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // separator7
            // 
            this.separator7.Name = "separator7";
            // 
            // separator8
            // 
            this.separator8.Name = "separator8";
            // 
            // separator3
            // 
            this.separator3.Name = "separator3";
            // 
            // separator11
            // 
            this.separator11.Name = "separator11";
            // 
            // separator10
            // 
            this.separator10.Name = "separator10";
            // 
            // separator9
            // 
            this.separator9.Name = "separator9";
            // 
            // separator4
            // 
            this.separator4.Name = "separator4";
            // 
            // grpReports
            // 
            this.grpReports.Items.Add(this.label5);
            this.grpReports.Items.Add(this.mnuReports);
            this.grpReports.Label = "Reports";
            this.grpReports.Name = "grpReports";
            // 
            // label5
            // 
            this.label5.Label = "   ";
            this.label5.Name = "label5";
            // 
            // grpDebug
            // 
            this.grpDebug.Items.Add(this.lblUrl);
            this.grpDebug.Items.Add(this.lblAliasID);
            this.grpDebug.Label = "Debug";
            this.grpDebug.Name = "grpDebug";
            // 
            // lblUrl
            // 
            this.lblUrl.Label = "Url";
            this.lblUrl.Name = "lblUrl";
            this.lblUrl.Text = null;
            // 
            // lblAliasID
            // 
            this.lblAliasID.Label = "AliasID";
            this.lblAliasID.Name = "lblAliasID";
            this.lblAliasID.Text = null;
            // 
            // grpVersion
            // 
            this.grpVersion.Items.Add(this.lblVersion);
            this.grpVersion.Items.Add(this.label6);
            this.grpVersion.Label = "Version";
            this.grpVersion.Name = "grpVersion";
            // 
            // lblVersion
            // 
            this.lblVersion.Label = "FAST BETA 5.11";
            this.lblVersion.Name = "lblVersion";
            // 
            // label6
            // 
            this.label6.Label = "09/25/2019";
            this.label6.Name = "label6";
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Image = global::Test_WorkBookOpen.Properties.Resources.Logo;
            this.button1.Label = "Amazon Devices";
            this.button1.Name = "button1";
            this.button1.ScreenTip = "Updates of FAST BETA 5.3";
            this.button1.ShowImage = true;
            this.button1.SuperTip = "Custom Addin,Dynamic Promo Operations,Included ALL option in Input Type for TCPU " +
    "view.";
            // 
            // btnMenuInitialize
            // 
            this.btnMenuInitialize.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnMenuInitialize.Dynamic = true;
            this.btnMenuInitialize.Image = global::Test_WorkBookOpen.Properties.Resources.excel1;
            this.btnMenuInitialize.Label = "Connect WorkBook";
            this.btnMenuInitialize.Name = "btnMenuInitialize";
            this.btnMenuInitialize.ShowImage = true;
            this.btnMenuInitialize.ItemsLoading += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnMenuInitialize_ItemsLoading);
            // 
            // btnDownloadData
            // 
            this.btnDownloadData.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnDownloadData.Image = global::Test_WorkBookOpen.Properties.Resources.download1;
            this.btnDownloadData.Label = "Download Data";
            this.btnDownloadData.Name = "btnDownloadData";
            this.btnDownloadData.ShowImage = true;
            this.btnDownloadData.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDownloadData_Click);
            // 
            // btnUploadData
            // 
            this.btnUploadData.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnUploadData.Image = global::Test_WorkBookOpen.Properties.Resources.upload1;
            this.btnUploadData.Label = "Upload Data";
            this.btnUploadData.Name = "btnUploadData";
            this.btnUploadData.ShowImage = true;
            this.btnUploadData.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUploadData_Click);
            // 
            // btnRefresh
            // 
            this.btnRefresh.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnRefresh.Image = global::Test_WorkBookOpen.Properties.Resources.refresh1;
            this.btnRefresh.Label = "Refresh Pivot";
            this.btnRefresh.Name = "btnRefresh";
            this.btnRefresh.ShowImage = true;
            this.btnRefresh.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRefresh_Click);
            // 
            // btnRefreshBransonData
            // 
            this.btnRefreshBransonData.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnRefreshBransonData.Image = global::Test_WorkBookOpen.Properties.Resources.refresh1;
            this.btnRefreshBransonData.Label = "Refresh Branson Data";
            this.btnRefreshBransonData.Name = "btnRefreshBransonData";
            this.btnRefreshBransonData.ShowImage = true;
            this.btnRefreshBransonData.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRefreshBransonData_Click);
            // 
            // btnContactSupport
            // 
            this.btnContactSupport.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnContactSupport.Enabled = false;
            this.btnContactSupport.Image = global::Test_WorkBookOpen.Properties.Resources.contact1;
            this.btnContactSupport.Label = "Contact Support";
            this.btnContactSupport.Name = "btnContactSupport";
            this.btnContactSupport.ShowImage = true;
            this.btnContactSupport.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnContactSupport_Click);
            // 
            // mnuReports
            // 
            this.mnuReports.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.mnuReports.Dynamic = true;
            this.mnuReports.Image = global::Test_WorkBookOpen.Properties.Resources.report1;
            this.mnuReports.Items.Add(this.menuVarianceReport);
            this.mnuReports.Items.Add(this.menuAuditReport);
            this.mnuReports.Items.Add(this.menuStatistics);
            this.mnuReports.Label = "Reports";
            this.mnuReports.Name = "mnuReports";
            this.mnuReports.ShowImage = true;
            // 
            // menuVarianceReport
            // 
            this.menuVarianceReport.Dynamic = true;
            this.menuVarianceReport.Image = global::Test_WorkBookOpen.Properties.Resources.Variance;
            this.menuVarianceReport.Label = "Variance Report";
            this.menuVarianceReport.Name = "menuVarianceReport";
            this.menuVarianceReport.ShowImage = true;
            this.menuVarianceReport.Visible = false;
            // 
            // menuAuditReport
            // 
            this.menuAuditReport.Dynamic = true;
            this.menuAuditReport.Enabled = false;
            this.menuAuditReport.Image = global::Test_WorkBookOpen.Properties.Resources.Audit;
            this.menuAuditReport.Label = "Audit Report";
            this.menuAuditReport.Name = "menuAuditReport";
            this.menuAuditReport.ShowImage = true;
            this.menuAuditReport.Visible = false;
            // 
            // menuStatistics
            // 
            this.menuStatistics.Dynamic = true;
            this.menuStatistics.Enabled = false;
            this.menuStatistics.Image = global::Test_WorkBookOpen.Properties.Resources.Statistics;
            this.menuStatistics.Label = "Statistics";
            this.menuStatistics.Name = "menuStatistics";
            this.menuStatistics.ShowImage = true;
            this.menuStatistics.Visible = false;
            this.menuStatistics.ItemsLoading += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.menuStatistics_ItemsLoading);
            // 
            // FAST
            // 
            this.Name = "FAST";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.grpInformation.ResumeLayout(false);
            this.grpInformation.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.grpPromoFilter.ResumeLayout(false);
            this.grpPromoFilter.PerformLayout();
            this.grpFilter.ResumeLayout(false);
            this.grpFilter.PerformLayout();
            this.grpOperations.ResumeLayout(false);
            this.grpOperations.PerformLayout();
            this.grpReports.ResumeLayout(false);
            this.grpReports.PerformLayout();
            this.grpDebug.ResumeLayout(false);
            this.grpDebug.PerformLayout();
            this.grpVersion.ResumeLayout(false);
            this.grpVersion.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu btnMenuInitialize;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpFilter;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ddnScenario;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ddnInputType;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ddnCurrency;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpOperations;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDownloadData;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUploadData;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRefresh;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnContactSupport;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpReports;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuVarianceReport;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpVersion;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lblVersion;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpInformation;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label2;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label3;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator7;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator8;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator9;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label5;
		internal Microsoft.Office.Tools.Ribbon.RibbonLabel label6;
		internal Microsoft.Office.Tools.Ribbon.RibbonMenu mnuReports;
		internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuAuditReport;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuStatistics;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ddnProductLine;
		internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator6;
		internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
		internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator5;
		internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ddnInterval;
		internal Microsoft.Office.Tools.Ribbon.RibbonLabel label9;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ddnCountry;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpPromoFilter;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label7;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ddnDeviceType;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator11;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator10;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRefreshBransonData;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpDebug;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox lblUrl;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox lblAliasID;
    }

    partial class ThisRibbonCollection
    {
        internal FAST Ribbon1
        {
            get { return this.GetRibbon<FAST>(); }
        }
    }
}
