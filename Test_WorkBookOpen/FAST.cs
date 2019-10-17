using System;
using System.Linq;
using Microsoft.Office.Tools.Ribbon;
using System.Data;
using System.Net;
using System.Configuration;
using System.IO;
using Test_WorkBookOpen.Classes;
using System.Windows.Forms;
using ExcelTool = Microsoft.Office.Tools.Excel;
using ExcelSheet = Microsoft.Office.Tools.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using System.Text.RegularExpressions;
using System.Drawing;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Deployment.Application;
using System.Security.Permissions;
using System.Security;
using System.Security.Policy;
using Microsoft.Office.Tools.Excel;
using System.Collections.Generic;

namespace Test_WorkBookOpen
{
    public partial class FAST
    {
        #region Variables Decleration   
        public static Excel.Worksheet _refSheetData, _productRevenue, _pivotProductRevenue, _reportProductRevenue, _reportProductAudit, _statistics, _pivotAuditReport, _promoInputTool, _vdp, _tcpu, _bransonPromotions;
        public static Excel.Range _checkRangeForUpload;
        public static int _dataSourceLength, labelItemArrayNumber, tagItemArrayNumber;
        public static byte _readOnly, _allFieldsRequiredConditionForUpload;
        public static DataSet _dsAllFilters, _dsDownloadData, _dsInitilaizeWorkbook, _dsVarianceReport, _dsAuditReport, _dsStatistics;
        public static bool IsFileOpend = false, _IsPromotionErrorHit = false;
        public static ExcelTool.Chart _pieChart = null, _columnChart = null, _coneColChart = null;
        public static DataSet _dsupdateVersion;
        // Variable Decleration for storing Dropdown Values in Text Format
        public static string _txtProcess,
                             _startRange, _endRange,
                             _userRole, _readOnlyStartMonth, _readOnlyEndMonth,
                             _downloadDataEditFile, _saveDataFile,
                             _scenarioStatus, _dataTypeValue, _varianceFlag, _variancePercent, _description, _varianceValue,
                             _promoCountryValue, _promoDeviceTypeValue, _promoCountryLabel, _promoDeviceLabel, _promoCountryRefreshTCPUVDPValue, _promodeviceRefreshTCPUVDPValue, _promoCountryRefreshTCPUVDPLable, _promodeviceRefreshTCPUVDPLabel,
                               _promoCountryBransonRefreshValue, _promodeviceBransonRefreshValue, _promoCountryBransonRefreshLable, _promodeviceBransonRefreshLabel;


        public static DateTime _startDateTimeStamp, _endDateTimeStamp;

        // Variable Decleration for storing Dropdown Id's in Number Format
        public static string _valueProcess, _valueScenario, _valueInputType, _valueCurrency, _valuePreviousScenario, _lastColumnName, _valueInterval, _valueProductLine;  // Ex: 1, 1, 1, 0
        public static ExcelTool.ListObject _loProductReveneue = null, _loProductRevenueReport = null, _loAuditReport = null, _loStatistics = null;
        public static bool _verifyDownloadForUpload;


        public static decimal _minValue, _maxValue;
        public static string _downloadedProcessValueForOfflineOnline, _downloadedScenarioValueForOfflineOnline,
                             _downloadedInputTypeValueForOfflineOnline, _downloadedCurrencyValueForOfflineOnline,
                            _downloadedIntervalValueForOfflineOnline, _downloadedProductLineValueForOfflineOnline,
                            _promoDownloadCountryValueForOfflineOnline, _promoDownloadDeviceTypeForOfflineOnline;


        public static string localPathDataTable = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);

        public static WebClient web;
        public static string baseUrl = Convert.ToString(ConfigurationManager.AppSettings["BaseURL"]);

        // public static string userName;
        public static string url, response;
        public static string userName = "anush";/*"gangipam";*//*"sunkr"*///;

        public FASTWebServiceAdapter _fastServiceAdapter = new FASTWebServiceAdapter();

        public ClsPromotions _promotions = new ClsPromotions();

        public static Timer timer = new Timer();
        public static Timer afterUninstalltimer = new Timer();

        public static int usedRows;
        //variables for auto updates
        public static string _aliad_id, _version;
        public static DateTime _lastUpdatedTimeStamp;

        //for offline enable/disable of buttons;

        public static bool isDownloadEnabled;
        public static bool isUploadEnabled;
        public static bool isTCPUVDPEnabled;
        public static bool isBransonEnabled;
        //Date 10_09_2019
        #endregion

        #region Ribbon Load

        /// <summary>
        /// This method will be called once our ribbon is loaded
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            btnDownloadData.Enabled = false;
            menuVarianceReport.Enabled = false;
            menuAuditReport.Enabled = false;
            btnRefreshBransonData.Visible = false;

            grpDebug.Visible = false;//2-phase change
            grpFilter.Visible = false;
            //for offline 
            updateControl();

            // Using the below to check for Updates
            // MessageBox.Show("before update in ribbon load");
            checkForUpdateManually();

            // Added for Saving all our XML Files in the FAST Folder
            Directory.CreateDirectory(localPathDataTable + "\\FAST");

            //userName = Convert.ToString(System.Security.Principal.WindowsIdentity.GetCurrent().Name).Split('\\')[1];
        }

        #endregion

        #region Check For Updates Manually
        /// <summary>
        /// This Method is used to check for the Updates manually
        /// </summary>
        private void checkForUpdateManually()
        {
            // var data = FASTWebServiceAdapter.updateFastVersionAutoUpdate("Anwesh", "Fast PV 5.2", _lastUpdatedTimeStamp);
            // MessageBox.Show("Before ApplcationDeployment");
            // MessageBox.Show();
            if (ApplicationDeployment.IsNetworkDeployed)
            {
                Cursor cur = Cursor.Current;
                Cursor.Current = Cursors.WaitCursor;

                try
                {
                    // Setup the trust level
                    var deployment = ApplicationDeployment.CurrentDeployment;
                    var appId = new ApplicationIdentity(deployment.UpdatedApplicationFullName);
                    var unrestrictedPerms = new PermissionSet(PermissionState.Unrestricted);
                    var appTrust = new ApplicationTrust(appId)
                    {
                        DefaultGrantSet = new PolicyStatement(unrestrictedPerms),
                        IsApplicationTrustedToRun = true,
                        Persist = true
                    };
                    ApplicationSecurityManager.UserApplicationTrusts.Add(appTrust);

                    var info = deployment.CheckForDetailedUpdate();

                    if (info.UpdateAvailable)
                    {
                        Cursor.Current = cur;
                        if (MessageBox.Show("An updated version is available. Would you like to update now?",
                            clsInformation.displayMessageTitle, MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            Cursor.Current = Cursors.WaitCursor;
                            deployment.Update();
                            Cursor.Current = cur;


                            System.Reflection.Assembly assemblyInfo = System.Reflection.Assembly.GetExecutingAssembly();
                            Uri uriCodeBase = new Uri(assemblyInfo.CodeBase);

                            string ClickOnceLocation = Path.GetDirectoryName(uriCodeBase.LocalPath.ToString());

                            string sub = ClickOnceLocation.Substring(0, (ClickOnceLocation.IndexOf("2.0") + 3));

                            _aliad_id = userName;
                            _version = lblVersion.Label;
                            _lastUpdatedTimeStamp = DateTime.Now;

                            FASTWebServiceAdapter.updateFastVersionAutoUpdate(_aliad_id, _version, _lastUpdatedTimeStamp);

                            displayAlerts(clsInformation.updateDownload, 3);

                            Directory.Delete(sub, true);
                            string[] dirslist = Directory.GetDirectories(sub);

                            foreach (string dirTodelete in dirslist)
                            {
                                Directory.Delete(dirTodelete, true);
                            }
                        }


                    }
                    else
                    {
                        Cursor.Current = cur;
                    }
                }
                catch (Exception ex)
                {
                    displayAlerts(ex.Message + "\n" + ex.InnerException != null ? ex.InnerException.ToString() : string.Empty, 1);
                }
                finally
                {
                    Cursor.Current = cur;
                }
            }
        }

        #endregion

        #region Dropdown Menu Click (Initialize Workbook on Click)

        /// <summary>
        /// Connect Workbook Menu Button on click this method will be called.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// 
        private void btnMenuInitialize_ItemsLoading(object sender, RibbonControlEventArgs e)
        {
            try
            {
                updateEvents(false);

                if (IsUninstall())
                {
                    return;
                }

                btnMenuInitialize.Items.Clear();

                // Added For Debugging Purpose
                if (lblAliasID.Text != "")
                    userName = lblAliasID.Text;
                else
                {
                    //userName = Convert.ToString(System.Security.Principal.WindowsIdentity.GetCurrent().Name).Split('\\')[1];
                    userName = "anush";
                }


                if (lblUrl.Text != "")
                    baseUrl = lblUrl.Text;
                else
                    baseUrl = Convert.ToString(ConfigurationManager.AppSettings["BaseURL"]);


                _dsInitilaizeWorkbook = FASTWebServiceAdapter.getProcessItems(userName);

                if (_dsInitilaizeWorkbook.Tables.Count > 0)
                {
                    for (int i = 0; i < _dsInitilaizeWorkbook.Tables[0].Rows.Count; i++)
                    {
                        RibbonButton rbtn = Globals.Factory.GetRibbonFactory().CreateRibbonButton();
                        rbtn.Label = Convert.ToString(_dsInitilaizeWorkbook.Tables[0].Rows[i].ItemArray[0]);
                        rbtn.Tag = Convert.ToString(_dsInitilaizeWorkbook.Tables[0].Rows[i].ItemArray[1]);
                        rbtn.Click += new RibbonControlEventHandler(initializeData);
                        btnMenuInitialize.Items.Add(rbtn);
                    }
                }
                else
                {
                    displayAlerts(clsInformation.userAccessMessage, 1);
                }
            }
            catch (Exception ex)
            {
                errorLog(ex.Message, "Add-in_btnMenuInitialize_ItemsLoading");
                handleAlerts(ex.Message);
            }
            finally
            {
                updateEvents(true);
            }
        }

        #endregion

        #region Generating Dropdown Filters

        /// <summary>
        /// Connect Workbook Ribbon Menu Item on click this method will be called to generate the dropdown Items
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void initializeData(object sender, RibbonControlEventArgs e)
            {
            try
            {
                updateEvents(false);

                if (IsUninstall())
                {
                    return;
                }



                #region Condition Check
                if (!checkConditions(null))
                {
                    return;
                }




                #endregion

                #region Other Process
                //

                _txtProcess = (sender as RibbonButton).Label;

                _valueProcess = getRequiredData((sender as RibbonButton).Label + "," + clsInformation.process);

                btnContactSupport.Enabled = true;

                ddnInterval.Enabled = false;
                ddnProductLine.Enabled = false;
                menuVarianceReport.Visible = false;

                DataTable table1 = new DataTable("tblButtons");

                if (_txtProcess == clsInformation.tcpuView)
                {
                    btnContactSupport.Enabled = true;

                    ddnCurrency.Visible = false;
                    ddnInterval.Enabled = true;
                    ddnInterval.Visible = true;
                    ddnProductLine.Enabled = true;
                    ddnProductLine.Visible = true;
                    ddnScenario.Visible = true;
                    ddnInputType.Visible = true;
                    grpReports.Visible = true;
                    grpPromoFilter.Visible = false;
                    grpFilter.Visible = true;
                    btnRefreshBransonData.Visible = false;
                    btnDownloadData.Visible = true;
                    btnUploadData.Visible = true;// recent requirements changed btnDownloadData is visible to true /
                    btnRefresh.Label = "Refresh Pivot";
                    btnRefresh.Enabled = true;

                    //anwesh
                    btnDownloadData.Enabled = true;
                    btnUploadData.Enabled = true;
                }
                else if (_txtProcess == clsInformation.accountingView)
                {
                    grpFilter.Visible = true;
                    btnContactSupport.Enabled = true;

                    ddnInterval.Visible = false;
                    ddnProductLine.Visible = false;
                    menuVarianceReport.Visible = false;
                    ddnCurrency.Visible = true;
                    ddnScenario.Visible = true;
                    ddnInputType.Visible = true;
                    grpReports.Visible = true;
                    grpPromoFilter.Visible = false;
                    grpFilter.Visible = true;
                    btnRefreshBransonData.Visible = false;
                    btnUploadData.Visible = true;
                    btnDownloadData.Visible = true; // recent requirements changed btnDownloadData is visible to true /
                    btnRefresh.Label = "Refresh Pivot";
                    //added by Anwesh
                    btnRefresh.Enabled = true;
                    btnDownloadData.Enabled = true;
                    btnUploadData.Enabled = true;
                }
                else if (string.Equals(_txtProcess, clsInformation.promotionsView, StringComparison.InvariantCultureIgnoreCase))
                {
                    //btnContactSupport.Enabled = true;
                    ddnCountry.Visible = true;
                    ddnDeviceType.Visible = true;
                    ddnCurrency.Visible = false;
                    ddnInterval.Visible = false;
                    ddnProductLine.Visible = false;
                    ddnScenario.Visible = false;
                    ddnInputType.Visible = false;
                    grpReports.Visible = false;
                    grpPromoFilter.Visible = true;
                    grpFilter.Visible = false;

                    btnRefresh.Label = "Refresh VDP/TCPU Data";
                    btnRefreshBransonData.Visible = true;

                }

                // Generating the Sheets using the sheet Names available in clsAmazonMasterSheetNames
                clsManageSheet.buildSheet(ref _refSheetData, clsInformation.referenceDataSheet);
                _refSheetData.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;

                _dsAllFilters = FASTWebServiceAdapter.getDropdownDataItems(userName, Convert.ToString((sender as RibbonButton).Tag));

                if (_dsAllFilters != null)
                {
                    if (_dsAllFilters.Tables.Count > 0)
                    {
                        
                        if (_dsAllFilters.Tables[0].TableName == clsInformation.userRole)
                        {
                            menuAuditReport.Visible = true;
                            menuStatistics.Visible = true;
                        }

                        if (_txtProcess == clsInformation.accountingView || _txtProcess == clsInformation.tcpuView)
                        {
                            clsManageSheet.buildSheet(ref _reportProductRevenue, clsInformation.productsRevenueReport);
                            clsManageSheet.buildSheet(ref _pivotProductRevenue, clsInformation.productsRevenuePivot);
                            clsManageSheet.buildSheet(ref _productRevenue, clsInformation.productRevenue);

                            clsDataSheet.buildSheet(_refSheetData);
                            _productRevenue.Select();
                        }




                        #region Audit Report 
                        if (_dsAllFilters.Tables.Count > 7)
                        {

                            if (_dsAllFilters.Tables[4].TableName == clsInformation.userRole)
                            {
                                int getUserRoleColumnCount = 0;

                                for (int i = 0; i < _dsAllFilters.Tables[4].Columns.Count; i++)
                                {
                                    if (Convert.ToString(_dsAllFilters.Tables[4].Columns[i]) == clsInformation.userRole)
                                    {
                                        getUserRoleColumnCount = i;
                                    }
                                }

                                if (_dsAllFilters.Tables[4].Rows.Count != 0)
                                {
                                    if (Convert.ToString(_dsAllFilters.Tables[4].Rows[0].ItemArray[getUserRoleColumnCount]).Contains(clsInformation.admin))
                                    {
                                        // For Audit Report
                                        if (Convert.ToString(_dsAllFilters.Tables[5].TableName) == clsInformation.auditScenario)
                                        {
                                            // Clearing the Menu Audit Report and making it Visible
                                            menuAuditReport.Items.Clear();
                                            menuAuditReport.Visible = true;
                                            menuAuditReport.Enabled = true;

                                            // Clearing the Menu Statistics Report and making it Visible
                                            menuStatistics.Items.Clear();
                                            menuStatistics.Visible = true;
                                            menuStatistics.Enabled = true;


                                            for (int i = 0; i < _dsAllFilters.Tables[5].Rows.Count; i++)
                                            {
                                                RibbonButton rbtn = Globals.Factory.GetRibbonFactory().CreateRibbonButton();
                                                rbtn.Label = Convert.ToString(_dsAllFilters.Tables[5].Rows[i].ItemArray[labelItemArrayNumber + 1]);
                                                rbtn.Tag = Convert.ToString(_dsAllFilters.Tables[5].Rows[i].ItemArray[tagItemArrayNumber]);
                                                rbtn.Click += new RibbonControlEventHandler(generateAuditReport);
                                                menuAuditReport.Items.Add(rbtn);

                                            }
                                            //Hiding Audit Report as it is not required but audit pivot is generated using this hidden sheet
                                            clsManageSheet.buildSheet(ref _reportProductAudit, clsInformation.productsAuditReport);
                                            if (_txtProcess == clsInformation.tcpuView)
                                            {
                                                _reportProductAudit.Visible = Excel.XlSheetVisibility.xlSheetHidden;

                                                clsManageSheet.buildSheet(ref _pivotAuditReport, clsInformation.productsAuditReportPivot);
                                                _pivotAuditReport.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                                                _pivotAuditReport.Select();
                                                moveSheets();
                                            }
                                            else
                                            {

                                                _reportProductAudit.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                                                if (_pivotAuditReport != null)
                                                {
                                                    string verifyPivotAuditReport = null;
                                                    ExcelTool.Workbook wrkbk = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook);

                                                    foreach (Excel.Worksheet sheet in wrkbk.Sheets)
                                                    {

                                                        if (sheet.Name == clsInformation.productsAuditReportPivot)
                                                        {
                                                            verifyPivotAuditReport = sheet.Name;
                                                            break;
                                                        }
                                                    }

                                                    if (verifyPivotAuditReport != null)
                                                    {
                                                        _pivotAuditReport.Visible = Excel.XlSheetVisibility.xlSheetVisible;

                                                        _pivotAuditReport.Select();
                                                        _pivotAuditReport.Visible = Excel.XlSheetVisibility.xlSheetHidden;

                                                    }

                                                }
                                                _reportProductAudit.Select();
                                                moveSheets();
                                            }

                                            clsManageSheet.buildSheet(ref _statistics, clsInformation.statistics);
                                            _statistics.Select();
                                            moveSheets();

                                        }
                                        else
                                        {
                                            menuAuditReport.Visible = false;
                                            menuStatistics.Visible = false;
                                            verifySheetsAndDelete();
                                        }

                                    }
                                    else
                                    {
                                        menuAuditReport.Visible = false;
                                        menuStatistics.Visible = false;
                                        verifySheetsAndDelete();
                                    }
                                }
                                else
                                {
                                    menuAuditReport.Visible = false;
                                    menuStatistics.Visible = false;
                                    verifySheetsAndDelete();
                                }
                            }
                            else
                            {
                                menuAuditReport.Visible = false;
                                menuStatistics.Visible = false;
                                verifySheetsAndDelete();
                            }

                        }

                        #endregion


                        //btnDownloadData.Enabled = true;

                        DeleteAllSheetsFromWorkBook();
                        //Promotions Related sheets creation and remove other sheets

                        if (string.Equals(_txtProcess, clsInformation.promotionsView, StringComparison.InvariantCultureIgnoreCase))
                        {
                            //Create PromotionTools Sheets

                            ExcelTool.Workbook excelWorkbook = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook);

                            List<string> sheetNames = new List<string>();
                            foreach (Excel.Worksheet sheet in excelWorkbook.Sheets)
                            {
                                sheetNames.Add(sheet.Name);

                                // for offline
                                if (sheet.Name == clsInformation.bransonPromotions)
                                {
                                   _bransonPromotions = sheet;
                                }
                                else if (sheet.Name == clsInformation.VDP)
                                {
                                    _vdp = sheet;
                                }
                                else if (sheet.Name == clsInformation.TCPU)
                                {
                                    _tcpu = sheet;
                                }
                                else if (sheet.Name == clsInformation.PROMO_INPUT_TOOL)
                                {
                                    _promoInputTool = sheet;
                                }
                                
                            }

                            if (!sheetNames.Contains(clsInformation.PROMO_INPUT_TOOL) && _promoInputTool == null)
                            {
                                clsManageSheet.buildSheet(ref _promoInputTool, clsInformation.PROMO_INPUT_TOOL);//2-phase change
                                _promoInputTool.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;
                            }

                            if (!sheetNames.Contains(clsInformation.TCPU) && _tcpu == null)
                            {
                                clsManageSheet.buildSheet(ref _tcpu, clsInformation.TCPU);
                                _tcpu.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;
                            }

                            if (!sheetNames.Contains(clsInformation.VDP) && _vdp == null)
                            {
                                clsManageSheet.buildSheet(ref _vdp, clsInformation.VDP);
                                _vdp.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;

                            }

                            if (!sheetNames.Contains(clsInformation.bransonPromotions) && _bransonPromotions == null)
                            {
                                clsManageSheet.buildSheet(ref _bransonPromotions, clsInformation.bransonPromotions);//2-phase change
                                _bransonPromotions.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;

                            }
                            

                            if (_bransonPromotions != null) _bransonPromotions.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;
                            if(_vdp != null) _vdp.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;
                            if (_tcpu != null) _tcpu.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;
                            if (_promoInputTool != null) _promoInputTool.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;

                            //DeleteAllSheetsFromWorkBook();

                            //Added by Anwesh

                            if (_dsAllFilters.Tables.Contains(clsInformation.promoOperations))
                            {
                                int tableIndex = _dsAllFilters.Tables.IndexOf(clsInformation.promoOperations);

                                foreach (DataRow row in _dsAllFilters.Tables[tableIndex].Rows)
                                {
                                    int columnIndex = _dsAllFilters.Tables[tableIndex].Rows.IndexOf(row);

                                    switch (row[1].ToString())
                                    {
                                        case clsInformation.downloadEnable:
                                            isDownloadEnabled = Convert.ToBoolean(_dsAllFilters.Tables[tableIndex].Rows[columnIndex][clsInformation.enable]);
                                            btnDownloadData.Enabled = isDownloadEnabled;
                                            if (isDownloadEnabled)
                                            {
                                                _promoInputTool.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                                            }
                                                
                                            break;
                                        case clsInformation.uploadEnable:
                                            isUploadEnabled = Convert.ToBoolean(_dsAllFilters.Tables[tableIndex].Rows[columnIndex][clsInformation.enable]);
                                            btnUploadData.Enabled = isUploadEnabled;
                                            break;
                                        case clsInformation.refreshTCPUVDPEnable:
                                            isTCPUVDPEnabled = Convert.ToBoolean(_dsAllFilters.Tables[tableIndex].Rows[columnIndex][clsInformation.enable]);
                                            btnRefresh.Enabled = isTCPUVDPEnabled;
                                            if (isTCPUVDPEnabled)
                                            {
                                                _vdp.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                                                _tcpu.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                                            }
                                            break;
                                        case clsInformation.refreshBransonEnable:
                                            isBransonEnabled = Convert.ToBoolean(_dsAllFilters.Tables[tableIndex].Rows[columnIndex][clsInformation.enable]);
                                            btnRefreshBransonData.Enabled = isBransonEnabled;
                                            if (isBransonEnabled) _bransonPromotions.Visible = Excel.XlSheetVisibility.xlSheetVisible; 
                                            break;
                                    }
                                }
                                
                            }

                            sheetNames.Clear();

                            foreach (Excel.Worksheet sheet in excelWorkbook.Sheets)
                            {
                                sheetNames.Add(sheet.Name);

                                // for offline
                                if (sheet.Name == clsInformation.bransonPromotions && sheet.Visible != Excel.XlSheetVisibility.xlSheetVeryHidden)
                                {
                                    sheet.Select();
                                    return;
                                }
                                else if (sheet.Name == clsInformation.VDP && sheet.Visible != Excel.XlSheetVisibility.xlSheetVeryHidden)
                                {
                                    sheet.Select();
                                    return;
                                }
                                else if (sheet.Name == clsInformation.TCPU && sheet.Visible != Excel.XlSheetVisibility.xlSheetVeryHidden)
                                {
                                    sheet.Select();
                                    return;
                                }
                                else if (sheet.Name == clsInformation.PROMO_INPUT_TOOL && sheet.Visible != Excel.XlSheetVisibility.xlSheetVeryHidden)
                                {
                                    sheet.Select();
                                    return;
                                }

                            }


                        }

                    }
                }



                //reArrangeSheets();

                //For temparary download
                //btnDownloadData.Enabled = true; / comment this code recent requirements /
                #endregion

            }
            catch (Exception ex)
            {
                errorLog(ex.Message, "Add-in_InitializeData");
                handleAlerts(ex.Message);
            }
            finally
            {
                if (checkConditions("finally"))
                {
                    ddnCountry.Items.Clear();
                    ddnDeviceType.Items.Clear();
                    ddnScenario.Items.Clear();
                    ddnInputType.Items.Clear();
                    ddnCurrency.Items.Clear();
                    ddnInterval.Items.Clear();
                    ddnProductLine.Items.Clear();
                    menuVarianceReport.Items.Clear();

                    if (checkConditions("finally"))
                    {
                        if (_txtProcess == clsInformation.promotions)
                        {
                            buildDropDownList(ddnCountry, clsInformation.defaultCountry, 1, 7);
                            buildDropDownList(ddnDeviceType, clsInformation.defaultDeviceType, 1, 8);
                        }
                        else
                        {
                            // Row Values and Column Values of Referencesheet is given here
                            buildDropDownList(ddnScenario, clsInformation.defaultScenario, 1, 0);
                            buildDropDownList(ddnCurrency, clsInformation.defaultCurrency, 1, 1);
                            buildDropDownList(ddnInputType, clsInformation.defaultInputType, 1, 2);
                            buildDropDownList(ddnInterval, clsInformation.defaultInterval, 1, 3);
                            buildDropDownList(ddnProductLine, clsInformation.defaultProductLine, 1, 6);
                        }

                    }
                }
                ddnCurrency.Enabled = false;

                updateEvents(true);
            }
        }

        

        private static void reArrangeSheets()
        {
            if (_txtProcess == clsInformation.accountingView || _txtProcess == clsInformation.tcpuView)
            {
                int totalSheets = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets.Count;

                if (_reportProductRevenue != null)
                    ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet).Move(
                         Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[totalSheets - 1]);

                if (_pivotProductRevenue != null)
                    ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet).Move(
                        Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[totalSheets - 2]);

                if (_productRevenue != null)
                    ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet).Move(
                        Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[totalSheets - 3]);
            }
        }

        private static void DeleteAllSheetsFromWorkBook()
        {
            //Delete all sheets
            ExcelTool.Workbook wrkbk = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook);
            Globals.ThisAddIn.Application.DisplayAlerts = false;

            switch (_txtProcess)
            {
                case clsInformation.accountingView:
                case clsInformation.tcpuView:
                    foreach (Excel.Worksheet sheet in wrkbk.Sheets)
                    {
                        if (sheet.Name == clsInformation.PROMO_INPUT_TOOL)
                        {
                            _promoInputTool = null;
                            sheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                            sheet.Delete();
                        }
                        else if (sheet.Name == clsInformation.TCPU)
                        {
                            _tcpu = null;
                            sheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                            sheet.Delete();
                        }
                        else if (sheet.Name == clsInformation.VDP)
                        {
                            _vdp = null;
                            sheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                            sheet.Delete();
                        }
                        else if (sheet.Name == clsInformation.referencePromo)
                        {   

                            sheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;
                            sheet.Delete();
                        }
                        else if (sheet.Name == clsInformation.bransonPromotions)
                        {
                            _bransonPromotions = null;
                            sheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                            sheet.Delete();
                        }
                    }
                    break;

                case clsInformation.promotions:
                    foreach (Excel.Worksheet sheet in wrkbk.Sheets)
                    {
                        if (sheet.Name == clsInformation.referenceDataSheet)
                        {
                            sheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;
                            sheet.Delete();
                        }
                        else if (sheet.Name == clsInformation.productRevenue)
                        {
                            sheet.Delete();
                        }
                        else if (sheet.Name == clsInformation.productsRevenuePivot)
                        {
                            sheet.Delete();
                        }
                        else if (sheet.Name == clsInformation.productsRevenueReport)
                        {
                            sheet.Delete();
                        }
                        else if (sheet.Name == clsInformation.productsAuditReportPivot)
                        {
                            sheet.Delete();
                        }
                        else if (sheet.Name == clsInformation.productsAuditReport)
                        {
                            sheet.Delete();
                        }
                        else if (sheet.Name == clsInformation.statistics)
                        {
                            sheet.Delete();
                        }
                        //else if (sheet.Name == clsInformation.PROMO_INPUT_TOOL && !isDownloadEnabled)
                        //{
                        //    sheet.Delete();
                        //}
                        //else if (sheet.Name == clsInformation.TCPU && !isTCPUVDPEnabled)
                        //{
                        //    sheet.Delete();
                        //}
                        //else if (sheet.Name == clsInformation.VDP && !isTCPUVDPEnabled)
                        //{
                        //    sheet.Delete();
                        //}
                        //else if (sheet.Name == clsInformation.bransonPromotions && !isBransonEnabled)
                        //{
                        //    sheet.Delete();
                        //}
                    }
                    break;
            }

            Globals.ThisAddIn.Application.DisplayAlerts = true;
        }


        /// <summary>
        /// This method is called to bind the data to the dropdown Items
        /// </summary>
        /// <param name="rddn">Ribbon Dropdown Item to which the data has to be binded</param>
        /// <param name="defaultText">Dropdown default Text</param>
        /// <param name="mode"></param>
        /// <param name="tableId">Id of the table to be used from dataset to bind the data</param>

        private void buildDropDownList(RibbonDropDown rddn, string defaultText, byte mode, byte tableId)
        {

            if (mode == 1)
            {
                if (rddn == null)
                    throw new Exception("Subsidiary drop down list can not be NULL when building it");

                rddn.Items.Clear();

                if (_dsAllFilters.Tables.Count > 0)
                {
                    if (tableId == 2 && _txtProcess == clsInformation.tcpuView)
                    {
                        DataRow dr = _dsAllFilters.Tables[2].NewRow();
                        dr["AllFieldsRequired"] = 0;
                        dr["Type"] = "InputType";
                        dr["Value"] = "ALL";
                        dr["MaximumValue"] = "1000.0";
                        //dr["MaximumValue"] = "50000.0";
                        //dr["MaximumValue"] = "50000.0";

                        dr["VarianceFlagType"] = "Greater Than";
                        dr["Id"] = 0;
                        dr["MinimumValue"] = "0";
                        //dr["MinimumValue"] = "-100.0";
                        dr["InputTypeDescription"] = "Description for ALL";
                        dr["InputDataType"] = "Decimal";
                        dr["VariancePercentage"] = "1.50";
                        dr["CurrencyConditionId"] = 2;
                        dr["IntervalId"] = "";
                        dr["Interval"] = "";
                        dr["isReadOnly"] = "";
                        dr["UserRole"] = "";
                        dr["VarianceValue"] = "0.250000";
                        _dsAllFilters.Tables[2].Rows.Add(dr);
                    }
                    for (int i = 0; i < _dsAllFilters.Tables[tableId].Rows.Count; i++)
                    {
                        RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                        item.Label = Convert.ToString(_dsAllFilters.Tables[tableId].Rows[i].ItemArray[labelItemArrayNumber + 1]);
                        item.Tag = Convert.ToInt32(_dsAllFilters.Tables[tableId].Rows[i].ItemArray[tagItemArrayNumber]);

                        rddn.Items.Add(item);
                    }
                }
            }

            if (!string.IsNullOrEmpty(defaultText))
            {
                RibbonDropDownItem defaultItem2 = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                defaultItem2.Label = defaultText;
                defaultItem2.Tag = 0;
                rddn.Items.Insert(0, defaultItem2);

                if (defaultText == clsInformation.defaultInterval)
                    if (rddn.Items.Count > 1)
                        rddn.SelectedItemIndex = 1;
            }
        }



        #endregion

        #region MoveSheets
        private void moveSheets()
        {
            int totalSheets = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets.Count;
            ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet).Move(
                Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[totalSheets]);
        }
        #endregion

        #region AuditReport

        private void generateAuditReport(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (IsUninstall())
                {
                    return;
                }
                #region Check Conditions
                if (!checkConditions("AuditScenario"))
                {
                    return;
                }

                #endregion

                string _valueAuditScenarioReport = Convert.ToString((sender as RibbonButton).Tag);

                // checking the condition for Downloading the data 
                if (_txtProcess != "" && _valueAuditScenarioReport != "" && _txtProcess != null && _valueAuditScenarioReport != null)
                {
                    // Setting all the events to disable here
                    updateEvents(false);

                    _valueProcess = getRequiredData(_txtProcess + "," + clsInformation.process);

                    // Building the Url
                    _dsAuditReport = FASTWebServiceAdapter.getAuditScenarioReport(userName, _valueProcess, _valueAuditScenarioReport);

                    if (_dsAuditReport.Tables.Count != 0 && _dsAuditReport.Tables.Count != 2)
                    {
                        clsManageSheet.buildAuditSheetBody(clsInformation.productsAuditReport, ref _loAuditReport, _dsAuditReport.Tables[1], _txtProcess, Convert.ToString((sender as RibbonButton).Label));
                        if (_reportProductAudit == null)
                            _reportProductAudit = (Excel.Worksheet)Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[clsInformation.productsRevenueReport]);

                        if (_txtProcess == clsInformation.tcpuView) //If tcpu view then build pivot for audit report
                        {
                            clsManageSheet.buildAuditReportPivotSheetBody(clsInformation.productsAuditReportPivot, _dsAuditReport.Tables[1], _txtProcess, Convert.ToString((sender as RibbonButton).Label));
                            _pivotAuditReport.Activate();
                        }
                        else
                            _reportProductAudit.Activate();

                        displayAlerts(clsInformation.reportSuccessfull, 1);

                    }
                    else
                    {
                        displayAlerts(clsInformation.noDataReport, 1);
                    }




                }
            }
            catch (Exception ex)
            {
                errorLog(ex.Message, "GenerateAuditReport");
                handleAlerts(ex.Message);
            }
            finally
            {
                updateEvents(true);
            }

        }



        private static void verifySheetsAndDelete()
        {
            ExcelTool.Workbook wrkbk = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook);

            foreach (Excel.Worksheet sheet in wrkbk.Sheets)
            {

                if (sheet.Name == clsInformation.productsAuditReport)
                {
                    Globals.ThisAddIn.Application.DisplayAlerts = false;
                    sheet.Delete();
                    Globals.ThisAddIn.Application.DisplayAlerts = true;
                    break;
                }
            }


            foreach (Excel.Worksheet sheet in wrkbk.Sheets)
            {
                if (sheet.Name != null)
                {
                    if (sheet.Name == clsInformation.statistics)
                    {
                        Globals.ThisAddIn.Application.DisplayAlerts = false;
                        sheet.Delete();
                        Globals.ThisAddIn.Application.DisplayAlerts = true;
                        break;
                    }
                }
            }

            foreach (Excel.Worksheet sheet in wrkbk.Sheets)
            {
                if (sheet.Name != null)
                {
                    if (sheet.Name == clsInformation.productsAuditReportPivot)
                    {
                        Globals.ThisAddIn.Application.DisplayAlerts = false;
                        sheet.Delete();
                        Globals.ThisAddIn.Application.DisplayAlerts = true;
                        break;
                    }
                }
            }
        }

        #endregion

        #region Inputtype ChangeEvent

        /// <summary>
        /// This method is called to bind the currency for the selected InputType and make the Dropdown to be disabled
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ddnInputType_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (IsUninstall())
                {
                    return;
                }

                _valueInputType = Convert.ToString(ddnInputType.SelectedItem.Tag);
                _downloadedInputTypeValueForOfflineOnline = _valueInputType;

                int id = (from DataRow dr in _dsAllFilters.Tables[2].Rows
                          where (string)dr["Value"] == Convert.ToString(ddnInputType.SelectedItem)
                          select Convert.ToInt32(dr["Id"])).FirstOrDefault();

                var currencyListForInputType = (from DataRow dr in _dsAllFilters.Tables[2].Rows
                                                where Convert.ToString(dr["Id"]) == Convert.ToString(id)
                                                select Convert.ToString(dr["CurrencyConditionId"]));

                if (currencyListForInputType != null && currencyListForInputType.Count() <= 1)
                {
                    if (currencyListForInputType.Count() > 0)
                    {
                        var currencyValue = (from DataRow dr in _dsAllFilters.Tables[1].Rows
                                             where Convert.ToString(dr["Id"]) == currencyListForInputType.FirstOrDefault()
                                             select Convert.ToString(dr["Value"]));

                        ddnCurrency.SelectedItem = ddnCurrency.Items.Where(x => x.Label.Equals(currencyValue.FirstOrDefault())).FirstOrDefault();

                    }
                    else
                        ddnCurrency.SelectedItemIndex = 0;

                    ddnCurrency.Enabled = false;
                }
                else
                    ddnCurrency.Enabled = true;
            }
            catch (Exception ex)
            {
                errorLog(ex.Message, "Add-in-ddnInputType_SelectionChanged");
                handleAlerts(ex.Message);
            }
        }
        #endregion

        #region Scenario Changed Event

        /// <summary>
        /// This method is used to bind the data to the Variance Menu Item on Scenario Changed
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ddnScenario_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (IsUninstall())
                {
                    return;
                }
                #region Condition Check

                if (!checkConditions("download"))
                {
                    return;
                }

                #endregion

                menuVarianceReport.Items.Clear();

                _valueScenario = Convert.ToString(ddnScenario.SelectedItem.Tag);

                if (Convert.ToString(ddnScenario.SelectedItem).Equals("--Select Scenario--"))
                {
                    menuVarianceReport.Items.Clear();
                    menuVarianceReport.Enabled = false;
                    menuStatistics.Enabled = false;


                }
                else
                {

                    DataSet dsVarianceReportMenu = new DataSet();
                    response = null;

                    _valueProcess = getRequiredData(_txtProcess + "," + clsInformation.process);

                    dsVarianceReportMenu = FASTWebServiceAdapter.getVarianceReportMenuItems(userName, _valueProcess, _valueScenario, _valueProductLine);

                    if (dsVarianceReportMenu.Tables.Count != 0)
                    {
                        menuVarianceReport.Visible = true;
                        for (int i = 0; i < dsVarianceReportMenu.Tables[0].Rows.Count; i++)
                        {
                            RibbonButton item = Globals.Factory.GetRibbonFactory().CreateRibbonButton();
                            item.Label = Convert.ToString(dsVarianceReportMenu.Tables[0].Rows[i].ItemArray[1]);
                            item.Tag = Convert.ToInt32(dsVarianceReportMenu.Tables[0].Rows[i].ItemArray[0]);
                            item.Click += new RibbonControlEventHandler(generateVarianceReport);
                            menuVarianceReport.Items.Add(item);
                        }
                        menuVarianceReport.Enabled = true;
                        mnuReports.Enabled = true;
                        menuStatistics.Enabled = true;
                    }
                    else
                    {
                        menuVarianceReport.Enabled = false;
                        menuVarianceReport.Visible = false;

                        if (_dsAllFilters.Tables.Count == 6)
                        {
                            if (_dsAllFilters.Tables[5].TableName == "AuditScenario")
                            {
                                if (_dsAllFilters.Tables[5].Rows.Count == 0)
                                {
                                    mnuReports.Enabled = false;
                                }
                            }
                        }
                    }
                }

                // interval displayed  
                if (_txtProcess == clsInformation.tcpuView)
                {
                    ddnInterval.Items.Clear();

                    for (int i = 0; i < _dsAllFilters.Tables[0].Rows.Count; i++)
                    {
                        DataRow dr = _dsAllFilters.Tables[0].Rows[i];

                        if (Convert.ToString(dr["Id"]) == Convert.ToString(ddnScenario.SelectedItem.Tag))
                        {
                            RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                            item.Label = Convert.ToString(dr["Interval"]);
                            item.Tag = Convert.ToString(dr["IntervalId"]);
                            ddnInterval.Items.Add(item);

                        }
                    }

                    RibbonDropDownItem defaultItem2 = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                    defaultItem2.Label = clsInformation.defaultInterval;
                    defaultItem2.Tag = 0;
                    ddnInterval.Items.Insert(0, defaultItem2);


                    if (ddnInterval.Items.Count > 1)
                        ddnInterval.SelectedItemIndex = 1;
                }
            }

            catch (Exception ex)
            {
                errorLog(ex.Message, "ddnScenario_SelectionChanged");
                handleAlerts(ex.Message);
            }
        }
        #endregion

        #region ProductLine Changed Event
        private void ddnProductLine_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            _valueProductLine = Convert.ToString(ddnProductLine.SelectedItem.Tag);
        }
        #endregion

        #region Interval

        /// <summary>
        /// This method is used to bind the data to the 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// 

        private void ddnInterval_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            _valueInterval = Convert.ToString(ddnInterval.SelectedItem.Tag);

        }
        #endregion

        #region Statistics
        private void menuStatistics_ItemsLoading(object sender, RibbonControlEventArgs e)
        {
            //if (IsUninstall())
            //{
            //    return;
            //}
            //#region Check Conditions
            //if (!checkConditions("AuditScenario"))
            //{
            //    return;
            //}

            //#endregion

            menuStatistics.Items.Clear();

            RibbonButton rbtn = Globals.Factory.GetRibbonFactory().CreateRibbonButton();
            rbtn.Label = Convert.ToString(clsInformation.statisticsScenarioLabel);
            rbtn.Tag = Convert.ToString(clsInformation.statisticsScenarioTag);
            rbtn.Click += new RibbonControlEventHandler(generateStatistics);
            menuStatistics.Items.Add(rbtn);


            rbtn = Globals.Factory.GetRibbonFactory().CreateRibbonButton();
            rbtn.Label = Convert.ToString(clsInformation.statisticsInputTypeLabel);
            rbtn.Tag = Convert.ToString(clsInformation.statisticsInputTypeTag);
            rbtn.Click += new RibbonControlEventHandler(generateStatistics);
            menuStatistics.Items.Add(rbtn);

            rbtn = Globals.Factory.GetRibbonFactory().CreateRibbonButton();
            rbtn.Label = Convert.ToString(clsInformation.statisticsUserLabel);
            rbtn.Tag = Convert.ToString(clsInformation.statisticsUserTag);
            rbtn.Click += new RibbonControlEventHandler(generateStatistics);
            menuStatistics.Items.Add(rbtn);
        }
        private void generateStatistics(object sender, RibbonControlEventArgs e)
        {
            try
            {
                updateEvents(false);

                if (_statistics != null)
                {
                    if (IsUninstall())
                    {
                        return;
                    }


                    #region Check Conditions
                    if (!checkConditions("AuditScenario"))
                    {
                        return;
                    }

                    #endregion

                    string selectedStatisticsItem = Convert.ToString((sender as RibbonButton).Tag);

                    if (ddnScenario.SelectedItem.ToString() == clsInformation.defaultScenario)
                    {
                        displayAlerts(clsInformation.scenarioDropdown, 1);
                        return;
                    }
                    _valueProcess = getRequiredData(_txtProcess + ",Process");
                    _valueScenario = Convert.ToString(ddnScenario.SelectedItem.Tag);

                    _dsStatistics = FASTWebServiceAdapter.getAuditScenarioReport(userName, _valueProcess, _valueScenario);

                    if (_dsStatistics.Tables.Count != 0 && _dsStatistics.Tables.Count != 2)
                    {

                        DataTable dt = new DataTable();

                        switch (selectedStatisticsItem)
                        {
                            case clsInformation.statisticsScenarioTag:
                                var getScenarioUploadData = (from DataRow dr in _dsStatistics.Tables[1].Rows
                                                             where (string)dr["UserAction"] == "Upload Data"
                                                             select Convert.ToString(dr["UserAction"])).ToList();

                                var getScenarioDownloadData = (from DataRow dr in _dsStatistics.Tables[1].Rows
                                                               where (string)dr["UserAction"] == "Download Data"
                                                               select Convert.ToString(dr["UserAction"])).ToList();

                                dt.Columns.Add("S.No");
                                dt.Columns.Add("Scenario");
                                dt.Columns.Add("Downloads");
                                dt.Columns.Add("Uploads");

                                dt.Rows.Add(1, Convert.ToString(ddnScenario.SelectedItem.Label), getScenarioDownloadData.Count, getScenarioUploadData.Count);



                                break;

                            case clsInformation.statisticsInputTypeTag:
                                // Getting all distinct InputTypes from the Datatable

                                DataView view = new DataView(_dsStatistics.Tables[1]);
                                DataTable distinctValues = new DataTable();
                                distinctValues = view.ToTable(true, "InputType");

                                dt.Columns.Add("S.No");
                                dt.Columns.Add("Input Type");
                                dt.Columns.Add("Downloads");
                                dt.Columns.Add("Uploads");

                                for (int t = 0; t < distinctValues.Rows.Count; t++)
                                {
                                    var getInputTypeDownloadDataActionCount = (from DataRow dr in _dsStatistics.Tables[1].Rows
                                                                               where (string)dr["UserAction"] == "Download Data" &&
                                                                               (string)dr["InputType"] == Convert.ToString(distinctValues.Rows[t].ItemArray[0])
                                                                               select Convert.ToString(dr["UserAction"])).ToList();

                                    var getInputTypeUploadDataActionCount = (from DataRow dr in _dsStatistics.Tables[1].Rows
                                                                             where (string)dr["UserAction"] == "Upload Data" &&
                                                                             (string)dr["InputType"] == Convert.ToString(distinctValues.Rows[t].ItemArray[0])
                                                                             select Convert.ToString(dr["UserAction"])).ToList();

                                    dt.Rows.Add(t + 1, Convert.ToString(distinctValues.Rows[t].ItemArray[0]), getInputTypeDownloadDataActionCount.Count, getInputTypeUploadDataActionCount.Count);


                                }

                                break;

                            case clsInformation.statisticsUserTag:


                                DataView view1 = new DataView(_dsStatistics.Tables[1]);
                                DataTable distinctValues1 = new DataTable();
                                distinctValues1 = view1.ToTable(true, "InputType", "UserName");



                                dt.Columns.Add("S.No");
                                dt.Columns.Add("User");
                                dt.Columns.Add("InputType");
                                dt.Columns.Add("Downloads");
                                dt.Columns.Add("Uploads");

                                for (int t = 0; t < distinctValues1.Rows.Count; t++)
                                {
                                    var getInputTypeDownloadDataActionCount = (from DataRow dr in _dsStatistics.Tables[1].Rows
                                                                               where (string)dr["UserAction"] == "Download Data" &&
                                                                               (string)dr["InputType"] == Convert.ToString(distinctValues1.Rows[t].ItemArray[0]) &&
                                                                               (string)dr["UserName"] == Convert.ToString(distinctValues1.Rows[t].ItemArray[1])
                                                                               select Convert.ToString(dr["UserName"])).ToList();

                                    var getInputTypeUploadDataActionCount = (from DataRow dr in _dsStatistics.Tables[1].Rows
                                                                             where (string)dr["UserAction"] == "Upload Data" &&
                                                                             (string)dr["InputType"] == Convert.ToString(distinctValues1.Rows[t].ItemArray[0]) &&
                                                                               (string)dr["UserName"] == Convert.ToString(distinctValues1.Rows[t].ItemArray[1])
                                                                             select Convert.ToString(dr["UserName"])).ToList();



                                    dt.Rows.Add(t + 1, Convert.ToString(distinctValues1.Rows[t].ItemArray[1]), Convert.ToString(distinctValues1.Rows[t].ItemArray[0]), getInputTypeDownloadDataActionCount.Count, getInputTypeUploadDataActionCount.Count);


                                }

                                break;
                        }

                        clsManageSheet.buildStatisticsSheetBody(clsInformation.statistics, ref _loStatistics, selectedStatisticsItem, _txtProcess, dt, ref _pieChart, ref _columnChart, ref _coneColChart, Convert.ToString((sender as RibbonButton).Label), Convert.ToString(ddnScenario.SelectedItem.Label));
                        displayAlerts(clsInformation.statisticsAlerts, 1);
                    }
                    else
                    {

                        displayAlerts(clsInformation.noDataStatistics, 1);
                    }
                }
                else
                {
                    displayAlerts(clsInformation.clickInitialize, 1);
                }
            }
            catch (Exception ex)
            {
                errorLog(ex.Message, "Add in-generateStatistics");
                handleAlerts(ex.Message);
            }
            finally
            {
                updateEvents(true);
            }

        }
        #endregion

        #region Download Data Functionality
        /// <summary>
        /// Download Button onclick this method will be called to perform the respected operations
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        private void btnDownloadData_Click(object sender, RibbonControlEventArgs e)
        {

            //testSheet();

            try
            {
                updateEvents(false);

                if (IsUninstall())
                {

                    updateEvents(true);
                    return;
                }

                if (string.Equals(_txtProcess, clsInformation.promotionsView, StringComparison.InvariantCultureIgnoreCase))
                {
                    // added by praveen
                    Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationManual;

                    _promoCountryValue = Convert.ToString(ddnCountry.SelectedItem.Tag);
                    _promoCountryLabel = Convert.ToString(ddnCountry.SelectedItem.Label);

                    _promoDeviceTypeValue = Convert.ToString(ddnDeviceType.SelectedItem.Tag);
                    _promoDeviceLabel = Convert.ToString(ddnDeviceType.SelectedItem.Label);


                    if (!productRevenueDownloadDataCheck())
                    {
                        updateEvents(true);

                        return;
                    }

                    if (!checkConditions("promo"))
                    {
                        updateEvents(true);
                        return;
                    }
                    _valueProcess = getRequiredData(_txtProcess + "," + clsInformation.process);
                    _promotions.promotionsView(userName, _valueProcess, ddnCountry.SelectedItem.Tag.ToString(), ddnDeviceType.SelectedItem.Tag.ToString(), _txtProcess, ddnCountry.SelectedItem.Label, ddnDeviceType.SelectedItem.Label);
                    if (_IsPromotionErrorHit != true)
                        displayAlerts(clsInformation.downloadSuccess1, 1);

                    _verifyDownloadForUpload = true;
                    IsFileOpend = true;

                    //added by praveen
                    Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic;

                    // Globals.ThisAddIn.Application.AutoCorrect.AutoFillFormulasInLists = false;
                    Globals.Ribbons.Ribbon1.Base.RibbonUI.ActivateTabMso(clsInformation.ribbonControlId);

                    updateEvents(true);
                    return;
                }


                #region Conditions
                if (!checkConditions("download"))
                {
                    return;
                }

                #endregion

                #region Creating FileNames
                string tmp = Regex.Replace(Convert.ToString(ddnInputType.SelectedItem).ToLower(), @"(\s+|@|&|'|\(|\)|<|>|#|:|-|\+|\?)", "");

                if (tmp.Length > 10)
                {
                    _saveDataFile = "\\FAST\\" + clsInformation.dataTableSave + (tmp.Substring(tmp.Length - 10)) + "_" + DateTime.Now.ToString("MM-dd-yyyy h:mm:ss tt").Replace(":", "-").Replace(" ", "-") + ".xml";
                }
                else if (tmp.Length > 5)
                {
                    _saveDataFile = "\\FAST\\" + clsInformation.dataTableSave + (tmp.Substring(tmp.Length - 5) + "_" + DateTime.Now.ToString("MM-dd-yyyy h:mm:ss tt").Replace(":", "-").Replace(" ", "-") + ".xml");
                }
                else
                {
                    _saveDataFile = "\\FAST\\" + clsInformation.dataTableSave + (tmp.Substring(0, tmp.Length) + "_" + DateTime.Now.ToString("MM-dd-yyyy h:mm:ss tt").Replace(":", "-").Replace(" ", "-") + ".xml");
                }

                #endregion

                #region DownloadProcess
                if (productRevenueDownloadDataCheck())
                {

                    // Getting ID's For the Selected Item Here
                    _valueProcess = getRequiredData(_txtProcess + "," + clsInformation.process);
                    _valueScenario = Convert.ToString(ddnScenario.SelectedItem.Tag);
                    _valueInputType = getRequiredData(Convert.ToString(ddnInputType.SelectedItem) + "," + clsInformation.inputType);
                    _valueCurrency = Convert.ToString(ddnCurrency.SelectedItem.Tag);
                    _valueInterval = ddnInterval.SelectedItem != null ? Convert.ToString(ddnInterval.SelectedItem.Tag) : "0";
                    _valueProductLine = ddnProductLine.SelectedItem != null ? Convert.ToString(ddnProductLine.SelectedItem.Tag) : "0";


                    // Storing the Initialize Menu, Dropdown selected Items for Upload purpose.
                    _downloadedProcessValueForOfflineOnline = _valueProcess;
                    _downloadedScenarioValueForOfflineOnline = _valueScenario;
                    _downloadedInputTypeValueForOfflineOnline = _valueInputType;
                    _downloadedCurrencyValueForOfflineOnline = _valueCurrency;
                    _downloadedIntervalValueForOfflineOnline = _valueInterval;
                    _downloadedProductLineValueForOfflineOnline = _valueProductLine;

                    _dsDownloadData = FASTWebServiceAdapter.getDownloadDataForUser(userName, _valueProcess, _valueScenario, _valueInputType, _valueCurrency, _valueInterval, _valueProductLine);

                    #region checking PreviousFiles Available
                    // Deleting the Previously Saved Offline File if Available

                    if (File.Exists(localPathDataTable + _saveDataFile))
                    {
                        File.Delete(localPathDataTable + _saveDataFile);
                    }
                    #endregion

                    // Checking whether two result sets are available or not.
                    if (_dsDownloadData.Tables.Count != 0)
                    {

                        _startRange = null;
                        _endRange = null;
                        _readOnlyStartMonth = null;
                        _readOnlyEndMonth = null;

                        // Chance of having only one result set that to the validation result set
                        if (_dsDownloadData.Tables[2].TableName != clsInformation.scenarioValidations)
                        {
                            downloadProcess(1, _txtProcess, Convert.ToString(ddnScenario.SelectedItem), Convert.ToString(ddnInputType.SelectedItem),
                                     Convert.ToString(ddnCurrency.SelectedItem), null, Convert.ToString(ddnInterval.SelectedItem), Convert.ToString(ddnProductLine.SelectedItem));


                            btnRefresh.Enabled = true;

                            updateEvents(true);
                            displayAlerts(clsInformation.downloadSuccess, 1);
                        }
                        else
                        {

                            downloadProcess(2, _txtProcess, Convert.ToString(ddnScenario.SelectedItem), Convert.ToString(ddnInputType.SelectedItem),
                            Convert.ToString(ddnCurrency.SelectedItem), null, Convert.ToString(ddnInterval.SelectedItem), Convert.ToString(ddnProductLine.SelectedItem));


                            btnRefresh.Enabled = true;

                            updateEvents(true);
                            displayAlerts(clsInformation.noDataDownload, 1);
                        }
                    }
                }
                #endregion
            }
            catch (Exception ex)
            {
                errorLog(ex.Message, "Add-btnDownloadData_click");
                updateEvents(true);
                handleAlerts(ex.Message);
            }

        }

        /// <summary>
        /// The Below method is called to perform the Operations like, building the sheetbody and pivot sheetbody
        /// </summary>
        /// <param name="mode"></param>
        public static void downloadProcess(byte mode, string process, string scenario, string inputtype, string currency, string previousscenario, string intervel, string productLine)//Added by Sita
        {
            switch (mode)
            {
                case 1:

                    clsDataSheet.addDataforOffline(_refSheetData);
                    clsManageSheet.buildSheetBody(clsInformation.productRevenue, ref _loProductReveneue, _dsDownloadData.Tables[2],
                                    process, scenario, inputtype, currency, previousscenario, intervel, productLine);
                    clsManageSheet.buildPivotSheetBody(clsInformation.productsRevenuePivot, _dsDownloadData.Tables[2]);


                    break;

                case 2:
                    // Generating only the Sheet Title and SheetBody with Headers but no data

                    clsManageSheet.buildSheetBodyHeader(clsInformation.productRevenue, ref _loProductReveneue, process, scenario, inputtype, currency, previousscenario, intervel, productLine);
                    clsManageSheet.buildPivotSheetBody(clsInformation.productsRevenuePivot, _dsDownloadData.Tables[0]);
                    break;
            }

            if (_productRevenue == null)
                clsManageSheet.buildSheet(ref _productRevenue, clsInformation.productRevenue);

            _productRevenue.Activate();


            // for Verifying data downloaded or not for Upload button click
            _verifyDownloadForUpload = true;
            IsFileOpend = true;


            Globals.Ribbons.Ribbon1.Base.RibbonUI.ActivateTabMso(clsInformation.ribbonControlId);

        }


        /// <summary>
        /// Download Button on click to check the validations this method is called
        /// </summary>
        /// <returns>True or False</returns>
        private bool productRevenueDownloadDataCheck()
        {

            switch (_txtProcess)
            {
                case clsInformation.accountingView:
                    if (Convert.ToString(ddnScenario.SelectedItem) == clsInformation.defaultScenario &&
                        Convert.ToString(ddnInputType.SelectedItem) == clsInformation.defaultInputType &&
                        Convert.ToString(ddnCurrency.SelectedItem) == clsInformation.defaultCurrency &&
                        Convert.ToString(ddnInterval.SelectedItem) == clsInformation.defaultIntervel)

                    {
                        displayAlerts(clsInformation.allDropdowns, 3);
                        return false;
                    }
                    else if (Convert.ToString(ddnScenario.SelectedItem) == clsInformation.defaultScenario &&
                             Convert.ToString(ddnInputType.SelectedItem) == clsInformation.defaultInputType)
                    {
                        displayAlerts(clsInformation.scenarioInputTypedropdown, 3);
                        return false;
                    }

                    else if (Convert.ToString(ddnScenario.SelectedItem) == clsInformation.defaultScenario &&
                              Convert.ToString(ddnCurrency.SelectedItem) == clsInformation.defaultCurrency)
                    {
                        displayAlerts(clsInformation.scenarioCurrencydropdown, 3);
                        return false;
                    }

                    else if (Convert.ToString(ddnInputType.SelectedItem) == clsInformation.defaultInputType &&
                              Convert.ToString(ddnCurrency.SelectedItem) == clsInformation.defaultCurrency)
                    {
                        displayAlerts(clsInformation.inputTypeCurrencyDropdown, 3);
                        return false;
                    }
                    else if (Convert.ToString(ddnInputType.SelectedItem) == clsInformation.defaultInputType &&
                              Convert.ToString(ddnInterval.SelectedItem) == clsInformation.defaultIntervel)
                    {
                        displayAlerts(clsInformation.inputTypeIntervelDropdown, 3);
                        return false;
                    }
                    else if (Convert.ToString(ddnScenario.SelectedItem) == clsInformation.defaultScenario)
                    {
                        displayAlerts(clsInformation.scenarioDropdown, 3);
                        return false;
                    }
                    else if (Convert.ToString(ddnInputType.SelectedItem) == clsInformation.defaultInputType)
                    {
                        displayAlerts(clsInformation.inputTypeDropdown, 3);
                        return false;
                    }
                    else if (Convert.ToString(ddnInterval.SelectedItem) == clsInformation.defaultIntervel)
                    {
                        displayAlerts(clsInformation.intervalDropdown, 3);
                        return false;
                    }

                    else if (ddnCurrency.Enabled && Convert.ToString(ddnCurrency.SelectedItem) == clsInformation.defaultCurrency)
                    {
                        displayAlerts(clsInformation.currencyType, 3);
                        return false;
                    }
                    break;


                case clsInformation.tcpuView:
                    if (Convert.ToString(ddnScenario.SelectedItem) == clsInformation.defaultScenario &&
                        Convert.ToString(ddnInterval.SelectedItem) == clsInformation.defaultInterval &&
                        Convert.ToString(ddnInputType.SelectedItem) == clsInformation.defaultInputType &&
                        //Convert.ToString(ddnInterval.SelectedItem) == clsInformation.defaultInterval &&
                        Convert.ToString(ddnProductLine.SelectedItem) == clsInformation.defaultProductLine)
                    {
                        displayAlerts(clsInformation.allDropdowns, 3);
                        return false;
                    }
                    else if (Convert.ToString(ddnScenario.SelectedItem) == clsInformation.defaultScenario &&
                          Convert.ToString(ddnInputType.SelectedItem) == clsInformation.defaultInputType &&
                          Convert.ToString(ddnInterval.SelectedItem) == clsInformation.defaultInterval &&
                          Convert.ToString(ddnProductLine.SelectedItem) == clsInformation.defaultProductLine)
                    {
                        displayAlerts(clsInformation.pleaseSelect + clsInformation.scenario + clsInformation.commaSeperator +
                           clsInformation.inputType + clsInformation.commaSeperator +
                           clsInformation.interval + clsInformation.andSeperator +
                           clsInformation.productLine + clsInformation.commaSeperator + clsInformation.dropdowns, 3);

                        return false;
                    }
                    else if (Convert.ToString(ddnScenario.SelectedItem) == clsInformation.defaultScenario &&
                          Convert.ToString(ddnInputType.SelectedItem) == clsInformation.defaultInputType &&
                        Convert.ToString(ddnInterval.SelectedItem) == clsInformation.defaultInterval)
                    {
                        displayAlerts(clsInformation.pleaseSelect + clsInformation.scenario + clsInformation.commaSeperator +
                         clsInformation.inputType + clsInformation.andSeperator +
                         clsInformation.interval + clsInformation.dropdowns, 3);

                        return false;
                    }
                    else if (Convert.ToString(ddnScenario.SelectedItem) == clsInformation.defaultScenario &&
                        Convert.ToString(ddnInputType.SelectedItem) == clsInformation.defaultInputType &&
                        Convert.ToString(ddnProductLine.SelectedItem) == clsInformation.defaultProductLine)
                    {
                        displayAlerts(clsInformation.pleaseSelect + clsInformation.scenario + clsInformation.commaSeperator +
                         clsInformation.inputType + clsInformation.andSeperator +
                         clsInformation.productLine + clsInformation.dropdowns, 3);

                        return false;
                    }
                    else if (Convert.ToString(ddnScenario.SelectedItem) == clsInformation.defaultScenario &&
                       Convert.ToString(ddnInterval.SelectedItem) == clsInformation.defaultIntervel &&
                       Convert.ToString(ddnProductLine.SelectedItem) == clsInformation.defaultProductLine)
                    {
                        displayAlerts(clsInformation.pleaseSelect + clsInformation.scenario + clsInformation.commaSeperator +
                         clsInformation.productLine + clsInformation.andSeperator +
                         clsInformation.interval + clsInformation.dropdowns, 3);

                        return false;
                    }
                    else if (Convert.ToString(ddnInputType.SelectedItem) == clsInformation.defaultInputType &&
                       Convert.ToString(ddnProductLine.SelectedItem) == clsInformation.defaultCurrency &&
                       Convert.ToString(ddnInterval.SelectedItem) == clsInformation.defaultInterval)
                    {
                        displayAlerts(clsInformation.pleaseSelect + clsInformation.inputType + clsInformation.commaSeperator +
                        clsInformation.productLine + clsInformation.andSeperator +
                        clsInformation.interval + clsInformation.dropdowns, 3);

                        return false;
                    }
                    else if (Convert.ToString(ddnScenario.SelectedItem) == clsInformation.defaultScenario &&
                      Convert.ToString(ddnProductLine.SelectedItem) == clsInformation.defaultProductLine &&
                      Convert.ToString(ddnInterval.SelectedItem) == clsInformation.defaultInterval)
                    {
                        displayAlerts(clsInformation.pleaseSelect + clsInformation.scenario + clsInformation.commaSeperator +
                        clsInformation.productLine + clsInformation.andSeperator +
                        clsInformation.interval + clsInformation.dropdowns, 3);
                        return false;
                    }

                    else if (Convert.ToString(ddnProductLine.SelectedItem) == clsInformation.defaultProductLine &&
                     Convert.ToString(ddnInterval.SelectedItem) == clsInformation.defaultInterval)
                    {
                        displayAlerts(clsInformation.pleaseSelect + clsInformation.interval + clsInformation.andSeperator +
                        clsInformation.productLine + clsInformation.dropdowns, 3);
                        return false;
                    }
                    else if (Convert.ToString(ddnScenario.SelectedItem) == clsInformation.defaultScenario &&
                         Convert.ToString(ddnProductLine.SelectedItem) == clsInformation.defaultProductLine)
                    {

                        displayAlerts(clsInformation.pleaseSelect + clsInformation.scenario + clsInformation.andSeperator +
                         clsInformation.productLine + clsInformation.dropdowns, 3);
                        return false;
                    }
                    else if (Convert.ToString(ddnScenario.SelectedItem) == clsInformation.defaultScenario &&
                         Convert.ToString(ddnInterval.SelectedItem) == clsInformation.defaultInterval)
                    {

                        displayAlerts(clsInformation.pleaseSelect + clsInformation.scenario + clsInformation.andSeperator +
                         clsInformation.interval + clsInformation.dropdowns, 3);
                        return false;
                    }
                    else if (Convert.ToString(ddnInputType.SelectedItem) == clsInformation.defaultInputType &&
                        Convert.ToString(ddnProductLine.SelectedItem) == clsInformation.defaultProductLine)
                    {

                        displayAlerts(clsInformation.pleaseSelect + clsInformation.inputType + clsInformation.andSeperator +
                         clsInformation.productLine + clsInformation.dropdowns, 3);
                        return false;
                    }
                    else if (Convert.ToString(ddnInterval.SelectedItem) == clsInformation.defaultInterval)
                    {

                        displayAlerts(clsInformation.intervalDropdown, 3);
                        return false;
                    }
                    else if (Convert.ToString(ddnProductLine.SelectedItem) == clsInformation.defaultProductLine)
                    {

                        displayAlerts(clsInformation.ProductLineDropdown, 3);
                        return false;
                    }
                    else if (Convert.ToString(ddnScenario.SelectedItem) == clsInformation.defaultScenario)
                    {
                        displayAlerts(clsInformation.scenarioDropdown, 3);
                        return false;
                    }
                    else if (Convert.ToString(ddnInputType.SelectedItem) == clsInformation.defaultInputType)
                    {
                        displayAlerts(clsInformation.inputTypeDropdown, 3);
                        return false;
                    }
                    break;

                case clsInformation.promotions:
                    if (ddnCountry.SelectedItem == null && ddnDeviceType.SelectedItem == null)
                    {
                        displayAlerts(clsInformation.clickInitialize, 3);
                        return false;
                    }

                    if (Convert.ToString(ddnCountry.SelectedItem) == clsInformation.defaultCountry &&
                      Convert.ToString(ddnDeviceType.SelectedItem) == clsInformation.defaultDeviceType)
                    {
                        displayAlerts(clsInformation.allDropdowns, 3);
                        return false;
                    }
                    else if (Convert.ToString(ddnCountry.SelectedItem) == clsInformation.defaultCountry)
                    {
                        displayAlerts(clsInformation.CountryDropdown, 3);
                        return false;
                    }
                    else if (Convert.ToString(ddnDeviceType.SelectedItem) == clsInformation.defaultDeviceType)
                    {
                        displayAlerts(clsInformation.DeviceTypeDropdown, 3);
                        return false;
                    }
                    break;
            }




            return true;
        }


        /// <summary>
        /// Download Data Button on click to get the required information this method is called
        /// </summary>
        /// <param name="Text">Text can be the Process or InputType</param>
        /// <returns></returns>
        public static string getRequiredData(string Text)
        {
            try
            {
                string[] rqtext = Text.Trim().Split(',');
                string IDs = "";

                if (Text.Contains(clsInformation.process))
                {
                    int id = (from DataRow dr in _dsInitilaizeWorkbook.Tables[0].Rows
                              where (string)dr["Value"] == rqtext[0]
                              select Convert.ToInt32(dr["Id"])).FirstOrDefault();

                    IDs = Convert.ToString(id);
                }
                else if (Text.Contains(clsInformation.inputType))
                {
                    int id = (from DataRow dr in _dsAllFilters.Tables[2].Rows
                              where (string)dr["Value"] == rqtext[0]
                              select Convert.ToInt32(dr["Id"])).FirstOrDefault();

                    _minValue = (from DataRow dr in _dsAllFilters.Tables[2].Rows
                                 where Convert.ToString(dr["Id"]) == Convert.ToString(id)
                                 select Convert.ToDecimal(Convert.ToDouble(dr["MinimumValue"]))).FirstOrDefault();


                    _maxValue = (from DataRow dr in _dsAllFilters.Tables[2].Rows
                                 where Convert.ToString(dr["Id"]) == Convert.ToString(id)
                                 select Convert.ToDecimal(Convert.ToDouble(dr["MaximumValue"]))).FirstOrDefault();

                    _allFieldsRequiredConditionForUpload = (from DataRow dr in _dsAllFilters.Tables[2].Rows
                                                            where Convert.ToString(dr["Id"]) == Convert.ToString(id)
                                                            select Convert.ToByte(dr["AllFieldsRequired"])).FirstOrDefault();

                    _dataTypeValue = (from DataRow dr in _dsAllFilters.Tables[2].Rows
                                      where Convert.ToString(dr["Id"]) == Convert.ToString(id)
                                      select Convert.ToString(dr["InputDataType"])).FirstOrDefault();

                    _varianceFlag = (from DataRow dr in _dsAllFilters.Tables[2].Rows
                                     where Convert.ToString(dr["Id"]) == Convert.ToString(id)
                                     select Convert.ToString(dr["VarianceFlagType"])).FirstOrDefault();

                    _variancePercent = (from DataRow dr in _dsAllFilters.Tables[2].Rows
                                        where Convert.ToString(dr["Id"]) == Convert.ToString(id)
                                        select Convert.ToString(dr["VariancePercentage"])).FirstOrDefault();


                    _description = (from DataRow dr in _dsAllFilters.Tables[2].Rows
                                    where Convert.ToString(dr["Value"]) == Convert.ToString(rqtext[0])
                                    select Convert.ToString(dr["InputTypeDescription"])).FirstOrDefault();

                    if (_txtProcess == clsInformation.tcpuView)
                    {
                        _varianceValue = (from DataRow dr in _dsAllFilters.Tables[2].Rows
                                          where Convert.ToString(dr["Value"]) == Convert.ToString(rqtext[0])
                                          select Convert.ToString(dr["Variancevalue"])).FirstOrDefault();


                    }


                    IDs = Convert.ToString(id);
                }
                return IDs;
            }
            catch (Exception ex)
            {
                errorLog(ex.Message, "getRequiredData");
                return "Error";
            }

        }

        #endregion

        #region Upload Data Functionality

        /// <summary>
        /// This method is called on ExportData Button click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>


        private void btnUploadData_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {

                if (IsUninstall())
                {
                    return;
                }

                updateEvents(false);

                //anwesh 24/08/2019
                ExcelTool.Workbook excelWorkbook = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook);

                List<string> sheetNames = new List<string>();
                foreach (Excel.Worksheet sheet in excelWorkbook.Sheets)
                {
                    sheetNames.Add(sheet.Name);

                    // for offline
                    if (sheet.Name == clsInformation.bransonPromotions)
                    {
                        _bransonPromotions = sheet;
                    }
                    else if (sheet.Name == clsInformation.VDP)
                    {
                        _vdp = sheet;
                    }
                    else if (sheet.Name == clsInformation.TCPU)
                    {
                        _tcpu = sheet;
                    }
                    else if (sheet.Name == clsInformation.PROMO_INPUT_TOOL)
                    {
                        _promoInputTool = sheet;
                    }

                }


                if (IsFileOpend == false)
                    fetchingRequiredDataForOffline();

                #region Promo Planning Upload
                if (_txtProcess == clsInformation.promotions)
                {
                    
                    ExcelTool.Worksheet verifyPromo = null;

                    if (sheetNames.Contains(clsInformation.PROMO_INPUT_TOOL) && _promoInputTool != null && _promoInputTool.Visible != Excel.XlSheetVisibility.xlSheetVeryHidden)
                    {
                        verifyPromo = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[clsInformation.PROMO_INPUT_TOOL]);
                    }
                    else
                    {
                        displayAlerts(clsInformation.noPromoInputTemplate, 4);
                        return;
                    }
                    
                    ExcelTool.Worksheet verifyVdp = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[clsInformation.VDP]);
                    ExcelTool.Worksheet verifyTcpu = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[clsInformation.TCPU]);


                    //ExcelTool.Worksheet verifyPromo = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[clsInformation.PROMO_INPUT_TOOL]);
                    //ExcelTool.Worksheet verifyVdp = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[clsInformation.VDP]);
                    //ExcelTool.Worksheet verifyTcpu = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[clsInformation.TCPU]);

                    if (verifyPromo != null || verifyVdp != null || verifyTcpu != null)
                    {
                        if (!checkConditions("promo"))
                        {
                            return;
                        }



                        _promoCountryValue = ddnCountry.SelectedItem != null ? Convert.ToString(ddnCountry.SelectedItem.Tag) : _promoCountryValue;
                        _promoDeviceTypeValue = ddnDeviceType.SelectedItem != null ? Convert.ToString(ddnDeviceType.SelectedItem.Tag) : _promoDeviceTypeValue;

                        //if (_promoCountryValue == "0" || _promoDeviceTypeValue == "0")
                        //{
                        if (!ClsPromotions.verifyDownloadForPromoUpload())
                        {
                            return;
                        }
                        //}


                        if (_promoCountryValue != _promoDownloadCountryValueForOfflineOnline ||
                            _promoDeviceTypeValue != _promoDownloadDeviceTypeForOfflineOnline)
                        {
                            displayAlerts(clsInformation.misMatchTypes, 2);
                            return;
                        }



                        string response = ClsPromotions.promotionPlanningUpload(userName, _valueProcess, _promoCountryValue, _promoDeviceTypeValue);
                        if (response != "")
                            if (response.Contains("UploadSuccess"))
                            {
                                displayAlerts(clsInformation.promotionsUploadSuccess, 1);
                            }
                            else
                            {
                                displayAlerts(response, 1);
                            }
                        return;
                    }
                }

                #endregion

                #region Conditions

                if (!checkConditions("upload"))
                {
                    return;
                }
                #endregion

                #region Export Data Code


                try
                {
                    if (!checkConditionsForUpload())
                    {
                        return;
                    }
                    if (!checkingValidationsForUpload())
                    {
                        return;
                    }
                    else
                    {

                        verifyChangesWhileUpload();

                    }

                }
                finally
                {
                    if (_productRevenue == null)
                        clsManageSheet.buildSheet(ref _productRevenue, clsInformation.productRevenue);
                    convertFromSheetNameToProtectAndUnProtect(1, _productRevenue.Name);


                    updateEvents(true);
                }


                #endregion

            }
            catch (Exception ex)
            {
                errorLog(ex.Message, "btnUploadData_Click");
                handleAlerts(ex.Message);
            }
            finally
            {
                updateEvents(true);
            }
        }


        /// <summary>
        /// Once all conditions are verified, then this method is called and the data will be sent for Upload
        /// </summary>
        private void verifyChangesWhileUpload()
        {
            if (_productRevenue == null)
                clsManageSheet.buildSheet(ref _productRevenue, clsInformation.productRevenue);


            try
            {

                DataSet dsUploadstatus = new DataSet();

                clsproductUpdateXMLManager.convertRangeToXml();


                //var serviceRequestUrl = string.Format("{0}/{1}?AliasId={2}&ProcessId={3}&ScenarioId={4}&InputTypeId={5}&CurrencyConditionId={6}&IntervalId={7}&View={8}&ProductLineId={9}",
                //                    FAST.baseUrl, "performanceUpload", "sunkr", Convert.ToInt32(_valueProcess), Convert.ToInt32(_valueScenario), Convert.ToInt32(_valueInputType), Convert.ToInt32(_valueCurrency), Convert.ToInt32(_valueInterval), FAST._txtProcess, _valueProductLine);

                //MessageBox.Show(serviceRequestUrl);

                //string stringXmlData = Convert.ToString(clsproductUpdateXMLManager._data);

                //MessageBox.Show(stringXmlData.Substring(0,100));

                //MessageBox.Show("Before Calling the Api For Account and TCPU");

                dsUploadstatus = FASTWebServiceAdapter.sendUploadDataForUser(userName, _valueProcess, _valueScenario, _valueInputType, _valueCurrency, _valueInterval, _valueProductLine);

                //MessageBox.Show("After Calling the Api");

                //System.Text.StringBuilder b = new System.Text.StringBuilder();
                //foreach (System.Data.DataRow r in dsUploadstatus.Tables[0].Rows)
                //{
                //    foreach (System.Data.DataColumn c in dsUploadstatus.Tables[0].Columns)
                //    {
                //        b.Append(c.ColumnName.ToString() + ":" + r[c.ColumnName].ToString());
                //    }
                //}
                //MessageBox.Show(b.ToString());

                //if data is returned from response 
                if (dsUploadstatus.Tables[0].Rows.Count > 0)
                {
                    // If Upload is successful,the below code will execute
                    if (Convert.ToString(dsUploadstatus.Tables[0].Rows[0].ItemArray[1]) == clsInformation.uploadSuccess)
                    {

                        ExcelTool.Worksheet inputTemplateSheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[clsInformation.productRevenue]);

                        inputTemplateSheet.Select();

                        string inputTypeValue = Convert.ToString(inputTemplateSheet.Cells[7, 3].Value);

                        removeBackGroundColorForBalnkCells();
                        displayAlerts("Inputs on " + inputTypeValue + " " + clsInformation.changesUploadSuccess, 1);



                        clsproductUpdateXMLManager.deleteFile();
                        clsproductUpdateXMLManager.clearXmlRoot();
                    }
                    else
                    {   
                        //MessageBox.Show(First else);
                        displayAlerts(clsInformation.uploadFail, 3);
                    }
                }
                else
                {
                    //MessageBox.Show(Second else);
                    displayAlerts(clsInformation.uploadFail, 3);
                }



            }
            catch (Exception ex)
            {
                errorLog(ex.Message, "Add-in-verifychangeswhileupload");
                handleAlerts(ex.Message);
            }

        }



        public static DataTable makeTableFromRange(Excel.Range headerRange, ExcelTool.Worksheet sheet, long usedRange)
        {
            DataTable table = new DataTable();
            int startingCol = _txtProcess == "Accounting View" ? 11 : 12;

            string[] months = new string[] { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" };

            table.Columns.Add("InputTemplateDataId");
            table.Columns.Add("ProductLineId");
            table.Columns.Add("WirelessId");
            table.Columns.Add("ProcessId");
            table.Columns.Add("MemoryId");
            table.Columns.Add("CurrencyId");
            table.Columns.Add("InputTypeId");
            table.Columns.Add("CountryId");
            table.Columns.Add("ProgramId");
            table.Columns.Add("ChannelId");
            table.Columns.Add("DTCPId");
            table.Columns.Add("ProcessName");
            table.Columns.Add("InputType");
            table.Columns.Add("Product Line");
            table.Columns.Add("Channel");
            table.Columns.Add("Country");
            table.Columns.Add("Program");
            table.Columns.Add("Memory");
            table.Columns.Add("Wireless");
            table.Columns.Add("DTCP");
            table.Columns.Add("Currency");


            DataTable newTable = _dsDownloadData.Tables[2].DefaultView.ToTable(false, "InputTemplateDataId",
                                                                                            "ProductLineId", "WirelessId", "ProcessId", "MemoryId", "CurrencyId", "InputTypeId", "CountryId", "ProgramId",
                                                                                            "ChannelId", "DTCPId",
                                                                                            "ProcessName", "InputType", "Product Line", "Channel", "Country", "Program", "Memory", "Wireless", "DTCP", "Currency");


            foreach (Excel.Range col in headerRange)
            {
                if (_txtProcess == "Accounting View")
                {
                    string[] abc = Convert.ToString(col.Text).Split('/');
                    int a = Convert.ToInt16(abc[0]) - 1;
                    string val = months[a] + "' " + abc[2].Substring(abc[2].Length - 2);
                    table.Columns.Add(val);
                }
                else
                {
                    table.Columns.Add(Convert.ToString(col.Text).Replace("'", "-"));
                }

            }

            // comment changed usedrange count
            for (long l = clsManageSheet.bodyRowStartingNumber + 1; l <= usedRange; l++)
            {

                DataRow row = table.NewRow();

                for (int i = 0; i < newTable.Columns.Count; i++)
                {
                    int rowId = Convert.ToInt32(l) - (clsManageSheet.bodyRowStartingNumber + 1);
                    row[i] = newTable.Rows[rowId].ItemArray[i];
                }

                int headerCount = headerRange.Count + startingCol;
                int nextColumnId = newTable.Columns.Count;

                for (int i = startingCol; i < headerCount; i++)
                {
                    row[nextColumnId] = sheet.Cells[l, i].Value;
                    nextColumnId++;
                }

                table.Rows.Add(row);



            }


            return table;
        }
        #endregion

        #region Refresh Pivot

        /// <summary>
        /// When the Refresh Pivot Button is clicked, this Method is called
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>


        private void btnRefresh_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                updateEvents(false);

                if (IsUninstall())
                {
                    return;
                }

                if ((sender as RibbonButton).Label == "Refresh Pivot")
                {
                    if (!checkConditions("download"))
                    {
                        return;
                    }


                    //if (_dsDownloadData != null)
                    //	usedRows = _dsDownloadData.Tables[2].Rows.Count + clsManageSheet.bodyRowStartingNumber;

                    ExcelTool.Workbook wrkbk = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook);

                    usedRows = 0;

                    foreach (Excel.Worksheet sheet in wrkbk.Sheets)
                    {
                        if (sheet.Name == clsInformation.productRevenue)
                        {
                            usedRows = sheet.UsedRange.Rows.Count;

                            break;
                        }
                    }

                    if (usedRows > clsManageSheet.bodyRowStartingNumber)
                        clsManageSheet.RefreshPivoteGrid(clsInformation.productsRevenuePivot, _dataSourceLength);
                    else
                        displayAlerts(clsInformation.refreshPivot, 1);
                }
                else if (((sender as RibbonButton).Label).ToUpper() == ("Refresh VDP/TCPU Data").ToUpper())
                {

                    // added by praveen
                    Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationManual;

                    //_promoCountryValue = Convert.ToString(ddnCountry.SelectedItem.Tag);
                    //_promoCountryLabel = Convert.ToString(ddnCountry.SelectedItem.Label);

                    //_promoDeviceTypeValue = Convert.ToString(ddnDeviceType.SelectedItem.Tag);
                    //_promoDeviceLabel = Convert.ToString(ddnDeviceType.SelectedItem.Label);


                    if (!checkConditions("PromoRefresh"))
                    {
                        return;
                    }

                    if (IsFileOpend == false)
                    {
                        fetchingRequiredDataForOffline();
                    }

                    _promoCountryRefreshTCPUVDPValue = ddnCountry.SelectedItem != null ? Convert.ToString(ddnCountry.SelectedItem.Tag) : _promoCountryRefreshTCPUVDPValue;
                    _promodeviceRefreshTCPUVDPValue = ddnDeviceType.SelectedItem != null ? Convert.ToString(ddnDeviceType.SelectedItem.Tag) : _promodeviceRefreshTCPUVDPValue;


                    _promoCountryRefreshTCPUVDPLable = ddnCountry.SelectedItem != null ? ddnCountry.SelectedItem.Label : _promoCountryRefreshTCPUVDPLable;
                    _promodeviceRefreshTCPUVDPLabel = ddnDeviceType.SelectedItem != null ? ddnDeviceType.SelectedItem.Label : _promodeviceRefreshTCPUVDPLabel;



                    if (_promoCountryRefreshTCPUVDPValue == "0" || _promodeviceRefreshTCPUVDPValue == "0")
                    {
                        if (!verifyDropdownsForPromo())
                        {
                            return;
                        }
                    }

                    _promotions.refreshVdpTcpu(userName, _valueProcess, _promoCountryRefreshTCPUVDPValue, _promodeviceRefreshTCPUVDPValue, _txtProcess, _promoCountryRefreshTCPUVDPLable, _promodeviceRefreshTCPUVDPLabel);

                    if (ClsPromotions.dtVdpRows != null && ClsPromotions.dtTcpuRows != null)
                        displayAlerts(clsInformation.refreshVdpTcpu, 1);

                    //anwesh 08/22/2019
                    _verifyDownloadForUpload = true;
                    Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
                }

                // else
                //                {
                //	displayAlerts(clsInformation.noDataReport, 1);
                //}

                Globals.Ribbons.Ribbon1.Base.RibbonUI.ActivateTabMso(clsInformation.ribbonControlId);

            }
            catch (Exception ex)
            {
                errorLog(ex.Message, (sender as RibbonButton).Label);
                handleAlerts(ex.Message);
            }
            finally
            {
                updateEvents(true);
            }
        }


        #endregion

        #region Refresh Branson
        private void btnRefreshBransonData_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                updateEvents(false);

                if (IsFileOpend == false)
                {
                    fetchingRequiredDataForOffline();
                }

                //anwesh refreshBranson
                if (!isDownloadEnabled && (_promoCountryValue == null || _promoDeviceTypeValue == null))
                {
                    _promoDownloadCountryValueForOfflineOnline = Convert.ToString(ddnCountry.SelectedItem.Tag);
                    _promoDownloadDeviceTypeForOfflineOnline = Convert.ToString(ddnDeviceType.SelectedItem.Tag);
                }

                _promoCountryBransonRefreshValue = ddnCountry.SelectedItem != null ? Convert.ToString(ddnCountry.SelectedItem.Tag) : _promoCountryBransonRefreshValue;
                _promodeviceBransonRefreshValue = ddnDeviceType.SelectedItem != null ? Convert.ToString(ddnDeviceType.SelectedItem.Tag) : _promodeviceBransonRefreshValue;


                _promoCountryBransonRefreshLable = ddnCountry.SelectedItem != null ? ddnCountry.SelectedItem.Label : _promoCountryBransonRefreshLable;
                _promodeviceBransonRefreshLabel = ddnDeviceType.SelectedItem != null ? ddnDeviceType.SelectedItem.Label : _promodeviceBransonRefreshLabel;

                //_promoCountryValue = Convert.ToString(ddnCountry.SelectedItem.Tag);
                //_promoCountryLabel = Convert.ToString(ddnCountry.SelectedItem.Label);

                //_promoDeviceTypeValue = Convert.ToString(ddnDeviceType.SelectedItem.Tag);
                //_promoDeviceLabel = Convert.ToString(ddnDeviceType.SelectedItem.Label);

                if (_promoCountryBransonRefreshValue == "0" || _promodeviceBransonRefreshValue == "0")
                {
                    if (!verifyDropdownsForPromo())
                    {
                        return;
                    }
                }

                _promotions.refreshBransonData(userName, _valueProcess, _promoCountryBransonRefreshValue, _promodeviceBransonRefreshValue, _txtProcess, _promoCountryBransonRefreshLable, _promodeviceBransonRefreshLabel);

                if (ClsPromotions.dtBransonPromotionsRows != null)
                    displayAlerts(clsInformation.refreshBransonData, 1);

                //anwesh 08/22/2019
                _verifyDownloadForUpload = true;

                Globals.Ribbons.Ribbon1.Base.RibbonUI.ActivateTabMso(clsInformation.ribbonControlId);
            }
            catch (Exception ex)
            {
                errorLog(ex.Message, (sender as RibbonButton).Label);
                handleAlerts(ex.Message);
            }
            finally
            {
                updateEvents(true);
            }
        }

        #endregion

        #region verifyDropdownsForPromo
        public bool verifyDropdownsForPromo()
        {
            if (ddnCountry.SelectedItem.Label == clsInformation.defaultCountry && ddnDeviceType.SelectedItem.Label == clsInformation.defaultDeviceType)
            {
                displayAlerts(clsInformation.CountryDeviceTypeDropdown, 1);
                return false;
            }

            if (ddnCountry.SelectedItem.Label == clsInformation.defaultCountry)
            {
                displayAlerts(clsInformation.CountryDropdown, 1);
                return false;
            }
            if (ddnDeviceType.SelectedItem.Label == clsInformation.defaultDeviceType)
            {
                displayAlerts(clsInformation.DeviceTypeDropdown, 1);
                return false;
            }

            return true;
        }
        #endregion

        #region ContactSupport

        /// <summary>
        /// Used to open Outlook easily for the clients to Contact the CustomerSupport
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnContactSupport_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (IsUninstall())
                {
                    return;
                }
                Outlook.Application objOutlook = new Outlook.Application();
                objOutlook.ActiveExplorer();

                Outlook.Application outlookApp = new Outlook.Application();

                Outlook.MailItem message = (Outlook.MailItem)outlookApp.CreateItem(Outlook.OlItemType.olMailItem);

                message.Subject = clsInformation.outlookSubject;

                if (_txtProcess == clsInformation.accountingView)
                {
                    message.Subject = clsInformation.outlookSubject;

                    message.Recipients.Add(clsInformation.outlookRecepients);
                }
                else if (_txtProcess == clsInformation.tcpuView)
                {
                    message.Subject = clsInformation.outlookSubject2;

                    message.Recipients.Add(clsInformation.outlookRecepients2);
                }
                else if (_txtProcess == clsInformation.promotionsView)
                {
                    message.Subject = clsInformation.outlookSubject3;

                    message.Recipients.Add(clsInformation.outlookRecepients3);
                }

                message.Body = "";

                message.Display(false);
            }
            catch (Exception ex)
            {
                errorLog(ex.Message, "Add-in_btnContactSupport_Click");
                handleAlerts(ex.Message);
            }
        }

        #endregion

        #region Variance Report
        /// <summary>
        /// When Variance Report menu Item on click, this method is called to generate the data
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        private void generateVarianceReport(object sender, RibbonControlEventArgs e)
        {

            try
            {
                if (IsUninstall())
                {
                    return;
                }
                #region Check Conditions
                if (!checkConditions("download"))
                {
                    return;
                }

                #endregion

                // checking the condition for Downloading the data 
                if (productRevenueDownloadDataCheck())
                {
                    // Setting all the events to disable here
                    updateEvents(false);

                    // Getting ID's For the Selected Item Here
                    string _valueProcessReport = getRequiredData(_txtProcess + "," + clsInformation.process);
                    string _valueScenarioReport = Convert.ToString(ddnScenario.SelectedItem.Tag);
                    string _valueInputTypeReport = getRequiredData(Convert.ToString(ddnInputType.SelectedItem) + "," + clsInformation.inputType);
                    string _valueCurrencyReport = Convert.ToString(ddnCurrency.SelectedItem.Tag);
                    string _valuePreviousScenarioReport = Convert.ToString((sender as RibbonButton).Tag);
                    string _valueIntervalReport = ddnInterval.SelectedItem != null ? Convert.ToString(ddnInterval.SelectedItem.Tag) : "0";
                    string _valueProductLineReport = ddnProductLine.SelectedItem != null ? Convert.ToString(ddnProductLine.SelectedItem.Tag) : "0";

                    // Building the Url
                    _dsVarianceReport = _fastServiceAdapter.GetVarianceReport(userName, _valueProcessReport, _valueScenarioReport, _valueInputTypeReport, _valueCurrencyReport, _valuePreviousScenarioReport, _valueIntervalReport, _valueProductLineReport);


                    clsManageSheet.buildVarianceReportBody(clsInformation.productsRevenueReport, ref _loProductRevenueReport, _dsVarianceReport.Tables[0], _txtProcess, Convert.ToString(ddnScenario.SelectedItem), Convert.ToString(ddnInputType.SelectedItem), Convert.ToString(ddnCurrency.SelectedItem), (sender as RibbonButton).Label, Convert.ToString(ddnInterval.SelectedItem), Convert.ToString(ddnProductLine.SelectedItem));
                    if (_reportProductRevenue == null)
                        _reportProductRevenue = (Excel.Worksheet)Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[clsInformation.productsRevenueReport]);

                    _reportProductRevenue.Activate();

                    displayAlerts(clsInformation.reportSuccessfull, 1);
                }
            }
            catch (Exception ex)
            {
                errorLog(ex.Message, "GenerateVarianceReport");
                handleAlerts(ex.Message);
            }
            finally
            {
                updateEvents(true);
            }
        }

        #endregion

        #region OFFLINE

        /// <summary>
        /// Before saving the Excel File, this method is called to save the changes done by the User for the data
        /// </summary>

        public static void savingRequiredDataForOffline()
        {
            try
            {
                if (IsUninstall())
                {
                    return;
                }



                // Added for FAST Folder, if it is not available, create again
                Directory.CreateDirectory(localPathDataTable + "\\FAST");


                // Added by Nihar For Promo Planning Offline
                if (_txtProcess == clsInformation.promotions)
                {
                        
                    savingPromoPlanningToolDataForOffline(isDownloadEnabled, isUploadEnabled, isTCPUVDPEnabled, isBransonEnabled);

                    return;
                }


                if (!checkConditions("Offline"))
                    return;

                if (_downloadedProcessValueForOfflineOnline != _valueProcess ||
                           _downloadedScenarioValueForOfflineOnline != _valueScenario ||
                           _downloadedInputTypeValueForOfflineOnline != _valueInputType ||
                           _downloadedCurrencyValueForOfflineOnline != _valueCurrency)
                {
                    _valueProcess = _downloadedProcessValueForOfflineOnline;
                    _valueScenario = _downloadedScenarioValueForOfflineOnline;
                    _valueInputType = _downloadedInputTypeValueForOfflineOnline;
                    _valueCurrency = _downloadedCurrencyValueForOfflineOnline;
                }

                // Adding (information) for Reference sheet to verify at the time of upload for offline functionality
                clsDataSheet.uploadButtonClickforOffline(_refSheetData);

                // Path for saving the DataTable XML File
                string filePath_For_DataTableData = localPathDataTable + _saveDataFile;

                //Converting the DataTable to XML and also saving it into the XML file along with the null values
                _dsDownloadData.WriteXml(filePath_For_DataTableData, XmlWriteMode.WriteSchema);

            }
            catch (Exception ex)
            {
                if (ex.Message != "No cells were found.")
                {
                    errorLog(ex.Message, "Saving_Required_Data_For_Offline");
                    handleAlerts(ex.Message);
                }

            }
        }


        private static void savingPromoPlanningToolDataForOffline(bool isDownloadEnabled, bool isUploadEnabled, bool isTCPUVDPEnabled, bool isBransonEnabled)
        {
            //commented by praveen1
            if (_promoDownloadCountryValueForOfflineOnline != _promoCountryValue ||
                _promoDownloadDeviceTypeForOfflineOnline != _promoDeviceTypeValue)
            {
                _promoCountryValue = _promoDownloadCountryValueForOfflineOnline;
                _promoDeviceTypeValue = _promoDownloadDeviceTypeForOfflineOnline;
            }

            ClsPromotions.uploadButtonClickforPromoOffline(isDownloadEnabled, isUploadEnabled, isTCPUVDPEnabled, isBransonEnabled);

        }

        /// <summary>
        /// This Method is called twice
        /// 1. Once the Excel File is saved to carry out the Operations of our application as in online for Offline
        /// 2. Once the Saved Excel sheet is reopened
        /// </summary>

        public static void fetchingRequiredDataForOffline()
        {

            try
            {

                //updateEvents(false);

                if (IsUninstall())
                {
                    return;
                }

                ExcelTool.Workbook wrkbk = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook);
                string _refSheet = null, promoReferenceSheet = null;
                foreach (Excel.Worksheet sheet in wrkbk.Sheets)
                {
                    if (sheet.Name == clsInformation.referenceDataSheet)
                    {
                        _refSheet = sheet.Name;
                    }

                    if (sheet.Name == clsInformation.referencePromo)
                        promoReferenceSheet = sheet.Name;
                }


                if (_refSheet != null)
                {
                    clsManageSheet.buildSheet(ref _refSheetData, _refSheet);


                    if (!checkConditions("Offline"))
                        return;

                    int getValueId = 13;

                    // For ProcessId, ScenarioId, InputTypeId, CurrencyConditionId
                    _valueProcess = Convert.ToString(_refSheetData.Cells[2, getValueId++].Value);
                    _valueScenario = Convert.ToString(_refSheetData.Cells[2, getValueId++].Value);
                    _valueInputType = Convert.ToString(_refSheetData.Cells[2, getValueId++].Value);
                    _valueCurrency = Convert.ToString(_refSheetData.Cells[2, getValueId++].Value);
                    _valueInterval = Convert.ToString(_refSheetData.Cells[2, getValueId++].Value);
                    _valueProductLine = Convert.ToString(_refSheetData.Cells[2, getValueId++].value);
                    _downloadedProcessValueForOfflineOnline = Convert.ToString(_refSheetData.Cells[2, getValueId++].Value);
                    _downloadedScenarioValueForOfflineOnline = Convert.ToString(_refSheetData.Cells[2, getValueId++].Value);
                    _downloadedInputTypeValueForOfflineOnline = Convert.ToString(_refSheetData.Cells[2, getValueId++].Value);
                    _downloadedCurrencyValueForOfflineOnline = Convert.ToString(_refSheetData.Cells[2, getValueId++].Value);
                    _downloadedIntervalValueForOfflineOnline = Convert.ToString(_refSheetData.Cells[2, getValueId++].Value);
                    _downloadedProductLineValueForOfflineOnline = Convert.ToString(_refSheetData.Cells[2, getValueId++].Value);
                    _minValue = Convert.ToDecimal(_refSheetData.Cells[2, getValueId++].Value);
                    _maxValue = Convert.ToDecimal(_refSheetData.Cells[2, getValueId++].Value);
                    _allFieldsRequiredConditionForUpload = Convert.ToByte(_refSheetData.Cells[2, getValueId++].Value);
                    _dataTypeValue = Convert.ToString(_refSheetData.Cells[2, getValueId++].Value);
                    _saveDataFile = Convert.ToString(_refSheetData.Cells[2, getValueId++].Value);
                    _startRange = Convert.ToString(_refSheetData.Cells[2, getValueId++].Value);
                    _endRange = Convert.ToString(_refSheetData.Cells[2, getValueId++].Value);
                    clsManageSheet.bodyRowStartingNumber = Convert.ToInt32(_refSheetData.Cells[2, getValueId++].Value);
                    clsManageSheet.formulaNextColumn = Convert.ToString(_refSheetData.Cells[2, getValueId++].Value);
                    //FAST.userName = Convert.ToString(_refSheetData.Cells[2, (getValueId+1)].Value);
                    //FAST.userName = Convert.ToString(System.Security.Principal.WindowsIdentity.GetCurrent().Name).Split('\\')[1];
                    userName = Convert.ToString(System.Security.Principal.WindowsIdentity.GetCurrent().Name).Split('\\')[1];


                    if (_txtProcess == null)
                        _txtProcess = Convert.ToString(_refSheetData.Cells[2, getValueId++].Value);

                    string filePathforXMLData = localPathDataTable + _saveDataFile;

                    if (filePathforXMLData != localPathDataTable)
                    {
                        // Creating a dataset object here and Reading the saved xml from the xml file
                        _dsDownloadData = new DataSet();
                        _dsDownloadData.ReadXml(filePathforXMLData);
                    }

                }

                if (promoReferenceSheet != null)
                {
                    ClsPromotions.promoUploadOnline();
                }
            }
            catch (Exception ex)
            {
                if (ex.Message != "No cells were found.")
                {
                    errorLog(ex.Message, "Fetching_Required_Data_For_Offline");
                    handleAlerts(ex.Message);
                }

            }
            finally
            {
                // updateEvents(true);
            }

        }

        #endregion

        #region ErrorLog

        /// <summary>
        /// Used to track all the Errors that are raised during the working of our application
        /// /// </summary>
        /// <param name="errormsg">What is the error that has occurred</param>
        /// <param name="errorpage">where the Error has occurred</param>
        /// 
        public static void errorLog(string errormsg, string errorpage)
        {
            try
            {
                web = new WebClient();

                response = null;
                url = string.Format("{0}/createErrorLog?AliasId={1}&ErrorCode={2}&ErrorMessage={3}&ErrorPage={4}&View={5}", baseUrl, userName, 0, errormsg, errorpage, FAST._txtProcess);
                response = web.DownloadString(url);

                return;
            }
            catch (Exception ex)
            {
                File.AppendAllText(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\FAST\\ErrorLog.txt", ex.Message + "Writing Error from Page -- " + errorpage + " -- " + DateTime.Now + Environment.NewLine);
                return;
            }

        }
        #endregion

        #region DisplayAlerts

        /// <summary>
        /// Used for displaying the Messages to the User
        /// </summary>
        /// <param name="displayMessage">Message to be displayed</param>
        /// <param name="messageType">And what type the Message is</param>
        public static void displayAlerts(string displayMessage, byte messageType)
        {
            if (messageType == 1)
            {
                MessageBox.Show(displayMessage, clsInformation.displayMessageTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (messageType == 2)
            {
                MessageBox.Show(displayMessage, clsInformation.displayMessageTitle, MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            else if (messageType == 3)
            {
                MessageBox.Show(displayMessage, clsInformation.displayMessageTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else if (messageType == 4) //anwesh 24/08/2019
            {
                MessageBox.Show(displayMessage, clsInformation.displayMessageTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

        }
        #endregion

        #region HandleAlerts

        /// <summary>
        /// Useful for Providing Customization Message to the user for errors
        /// </summary>
        /// <param name="message"></param>
        public static void handleAlerts(string message)
        {
            if (message.Contains("Could not find file"))
            {
                displayAlerts(clsInformation.offlineIssue, 1);
            }
            else
            {
                switch (message)
                {
                    case "The request was aborted: The operation has timed out.":
                    case "The remote name could not be resolved":
                        displayAlerts(clsInformation.corpNetworkIssue, 1);
                        break;

                    case "Bad Gateway.":
                    case "meta":
                    case "The remote server returned an error: (500) Internal Server Error.":
                    case "Unable to connect to the remote server":
                        displayAlerts(clsInformation.normalNetworkIssue, 1);
                        break;

                    case "Exception from HRESULT: 0x800A03EC":
                    case "Exception from HRESULT: 0x800401A8":
                        displayAlerts(clsInformation.multipleWorkbooks, 1);
                        break;

                    case "The underlying connection was closed: The connection was closed unexpectedly.":
                        displayAlerts(clsInformation.multipleWorkbooksNewly, 1);
                        break;

                    case "The operation has timed out":
                    case "The remote server returned an error: (502) Bad Gateway.":
                    case "The remote server returned an error: (404) Not Found.":
                        displayAlerts(clsInformation.dataNotUploaded, 1);
                        break;

                    case "Operation aborted (Exception from HRESULT: 0x80004004 (E_ABORT))":
                        displayAlerts(clsInformation.contactSupport, 1);
                        break;

                    case "The remote server returned an error: (503) Server Unavailable.":
                        displayAlerts(clsInformation.serviceUnavailable, 1);
                        break;

                    default:
                        displayAlerts(clsInformation.errorMessage, 3);
                        break;
                }
            }

        }
        #endregion

        #region EnableEvents

        /// <summary>
        /// Use for Improving the Performance of the Application
        /// </summary>
        /// <param name="status"> Specidfies true or false</param>

        public static void updateEvents(bool status)
        {
            Globals.ThisAddIn.Application.EnableEvents = status;
            Globals.ThisAddIn.Application.ScreenUpdating = status;
        }
        #endregion

        #region CheckConditions
        /// <summary>
        /// Checks the Conditions that are reqiured for the respected Operations
        /// </summary>
        /// <param name="conditioncheck"> for Upload or download or etc..</param>
        /// <returns>true or false</returns>
        public static bool checkConditions(string conditioncheck)
        {
            #region 
            //if (conditioncheck != "Offline")
            //{
            //    //For verifying whether any cell is in edit mode or not.
            //    if (isCellBeingEdited(Globals.ThisAddIn.Application))
            //    {
            //        if (conditioncheck != "finally")
            //            displayAlerts(clsInformation.editMode, 1);
            //        return false;
            //    }
            //}

            //// verifying whether there is any active workbook available or not in the Excel Window
            //if (Globals.ThisAddIn.Application.ActiveWorkbook == null)
            //{
            //    // showing a message box as there is no active workbook available
            //    displayAlerts(clsInformation.noActiveWorkbook, 1);
            //    return false;
            //}

            //if (conditioncheck == "promo" || conditioncheck == "PromoRefresh")
            //{

            //    ExcelTool.Worksheet verifyPromo = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[clsInformation.PROMO_INPUT_TOOL]);
            //    ExcelTool.Worksheet verifyVdp = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[clsInformation.VDP]);
            //    ExcelTool.Worksheet verifyTcpu = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[clsInformation.TCPU]);

            //    if (verifyPromo == null || verifyVdp == null || verifyTcpu == null)
            //    {
            //        displayAlerts(clsInformation.clickInitialize, 1);

            //        return false;
            //    }


            //    if (conditioncheck == "PromoRefresh")
            //    {
            //        if (verifyPromo.UsedRange.Rows.Count < 11)// if (verifyVdp.UsedRange.Rows.Count < 11 || verifyTcpu.UsedRange.Rows.Count < 11)
            //        {
            //            displayAlerts(clsInformation.refreshVdpTcpuCheck, 1);
            //            return false;
            //        }
            //    }


            //}


            //// The below condition will be checked for both the Upload and Download Processes
            //if (conditioncheck == "download" || conditioncheck == "upload" || conditioncheck == "Offline")
            //{
            //    ExcelTool.Workbook wrkbk = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook);

            //    string refWorksheet = null, inputTemplateSheet = null, varianceReport = null, inputTemplatePivot = null;

            //    foreach (Excel.Worksheet sheet in wrkbk.Sheets)
            //    {
            //        if (sheet.Name == clsInformation.referenceDataSheet)
            //        {
            //            refWorksheet = clsInformation.referenceDataSheet;
            //        }
            //        else if (sheet.Name == clsInformation.productRevenue)
            //        {
            //            inputTemplateSheet = clsInformation.productRevenue;
            //        }
            //        else if (sheet.Name == clsInformation.productsRevenuePivot)
            //        {
            //            inputTemplatePivot = clsInformation.productsRevenuePivot;
            //        }
            //        else if (sheet.Name == clsInformation.productsRevenueReport)
            //        {
            //            varianceReport = clsInformation.productsRevenueReport;
            //        }
            //    }

            //    //The below check is for Offline only
            //    if (conditioncheck == "Offline")
            //    {
            //        if (refWorksheet != null && (inputTemplateSheet == null || varianceReport == null || inputTemplatePivot == null))
            //        {
            //            displayAlerts(clsInformation.clickInitialize, 1);

            //            return false;
            //        }
            //        else if (refWorksheet == null)
            //            return false;
            //    }

            //    // Below is for Download and upload basic check

            //    if (refWorksheet == null || inputTemplateSheet == null || varianceReport == null || inputTemplatePivot == null)
            //    {

            //        //ExcelTool.Worksheet verifyPromoInputTool = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[clsInformation.PROMO_INPUT_TOOL]);
            //        //ExcelTool.Worksheet verifyVdp = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[clsInformation.VDP]);
            //        //ExcelTool.Worksheet verifyTcpu = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[clsInformation.TCPU]);

            //        string verifyPromoInputTool = null, verifyVdp = null, verifyTcpu = null;

            //        foreach (Excel.Worksheet sheet in wrkbk.Sheets)
            //        {
            //            if (sheet.Name == clsInformation.PROMO_INPUT_TOOL)
            //            {
            //                verifyPromoInputTool = clsInformation.PROMO_INPUT_TOOL;
            //            }
            //            else if (sheet.Name == clsInformation.VDP)
            //            {
            //                verifyVdp = clsInformation.VDP;
            //            }
            //            else if (sheet.Name == clsInformation.TCPU)
            //            {
            //                verifyTcpu = clsInformation.TCPU;
            //            }
            //        }

            //        if (verifyPromoInputTool == null || verifyVdp == null || verifyTcpu == null)
            //        {
            //            displayAlerts(clsInformation.clickInitialize, 1);

            //            return false;
            //        }


            //    }

            //    // The below Condition will be checked only for Upload Process
            //    if (conditioncheck == "upload")
            //    {
            //        if (!checkConditionsForUpload())
            //        {
            //            return false;
            //        }
            //    }



            //    if (conditioncheck == "AuditScenario")
            //    {
            //        ExcelTool.Worksheet verifyAuditReport = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[clsInformation.productsAuditReport]);


            //        if (verifyAuditReport == null)
            //        {
            //            displayAlerts(clsInformation.auditReportSheetNull, 1);
            //            return false;
            //        }

            //    }

            //}


            //return true;
            #endregion

            if (IsUninstall())
            {
                return false;
            }


            // verifying whether there is any active workbook available or not in the Excel Window
            if (Globals.ThisAddIn.Application.ActiveWorkbook == null)
            {
                // showing a message box as there is no active workbook available
                FAST.displayAlerts(clsInformation.noActiveWorkbook, 1);
                return false;
            }

            ExcelTool.Workbook wrkbk = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook);
            string inputTemplateSheet = null, promoInputToolSheet = null;

            switch (conditioncheck)
            {
                case "download":
                case "Save":
                    if (isCellBeingEdited(Globals.ThisAddIn.Application))
                    {
                        if (conditioncheck != "Save")
                        {
                            FAST.displayAlerts(clsInformation.editMode, 1);
                            return false;
                        }
                        return true;
                    }

                    break;
                case "upload":
                case "Offline":

                    foreach (Excel.Worksheet sheet in wrkbk.Sheets)
                    {
                        switch (sheet.Name)
                        {
                            case clsInformation.productRevenue:
                                inputTemplateSheet = clsInformation.productRevenue;
                                break;
                            case clsInformation.PROMO_INPUT_TOOL:
                                promoInputToolSheet = clsInformation.PROMO_INPUT_TOOL;
                                break;
                        }
                    }
                    //added for offline excel save sheets
                    if (inputTemplateSheet == clsInformation.productRevenue || promoInputToolSheet == clsInformation.PROMO_INPUT_TOOL)
                    {
                        switch (_txtProcess)
                        {
                            case clsInformation.accountingView:
                            case clsInformation.tcpuView:
                                if (inputTemplateSheet == null)
                                {
                                    FAST.displayAlerts(clsInformation.clickInitialize, 1);
                                    return false;
                                }
                                break;
                            case clsInformation.promotions:
                                if (promoInputToolSheet == null)
                                {
                                    FAST.displayAlerts(clsInformation.clickInitialize, 1);
                                    return false;
                                }
                                break;
                        }
                    }
                    else
                        return false;

                    // Online upload click with out downloading the data
                    if (inputTemplateSheet == null && promoInputToolSheet == null)
                    {
                        FAST.displayAlerts(clsInformation.clickInitialize, 1);
                        return false;
                    }

                    break;

                case "Empty":

                    foreach (Excel.Worksheet sheet in wrkbk.Sheets)
                    {
                        switch (sheet.Name)
                        {
                            case clsInformation.PROMO_INPUT_TOOL:
                                promoInputToolSheet = clsInformation.PROMO_INPUT_TOOL;
                                break;
                        }
                    }

                    if (promoInputToolSheet == null)
                        return false;
                    break;
            }
            return true;
        }
        #endregion

        #region Excel Edit Mode

        /// <summary>
        /// TO Identify whether the Excel is being edited or not to perform any operation
        /// </summary>
        /// <param name="excelApp"></param>
        /// <returns></returns>
        private static bool isCellBeingEdited(Excel.Application excelApp)
        {
            CommandBarControl cbc = excelApp.CommandBars.FindControl(1, 18, Type.Missing, Type.Missing);
            return cbc != null && !cbc.Enabled;
        }
        #endregion

        #region checkConditionsForUpload
        /// <summary>
        /// Conditions that are to be checked for Upload
        /// </summary>
        /// <returns>true or false</returns>
        private static bool checkConditionsForUpload()
        {
            if ((_txtProcess != clsInformation.accountingView && _txtProcess != clsInformation.tcpuView && _txtProcess != null))
            {
                return true;
            }


            ExcelTool.Worksheet inputTemplateSheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[clsInformation.productRevenue]);

            if (_dsDownloadData != null)
                // for veriying usedrange count
                usedRows = _dsDownloadData.Tables[2].Rows.Count + 14;
            else
            {
                usedRows = inputTemplateSheet.UsedRange.Rows.Count;

                //if (IsFileOpend != true)
                //{

                // IsFileOpend = true;
                // }

            }

            if (usedRows < 15)
            {
                displayAlerts(clsInformation.verifyDownloadforUpload, 1);
                return false;
            }

            string checkValue = _scenarioStatus != null ? _scenarioStatus.ToLower() : "Empty";
            if (checkValue == "close" && _userRole == "User")
            {
                displayAlerts(Convert.ToString(inputTemplateSheet.Cells[6, 3].Value) + clsInformation.scenarioClose, 2);

                return false;
            }

            if (_downloadedProcessValueForOfflineOnline != _valueProcess ||
                _downloadedScenarioValueForOfflineOnline != _valueScenario ||
                _downloadedInputTypeValueForOfflineOnline != _valueInputType ||
                _downloadedCurrencyValueForOfflineOnline != _valueCurrency ||
                 _downloadedIntervalValueForOfflineOnline != _valueInterval ||
                _downloadedProductLineValueForOfflineOnline != _valueProductLine)
            {
                displayAlerts(clsInformation.misMatchTypes, 1);

                return false;
            }

            // Added mainly to Verify the downloaded Input Type and Uploaded Input Type
            string downloadedInputType = Convert.ToString(inputTemplateSheet.Cells[7, 3].Value);
            int downloadedInputTypeId = getInputTypeIdForUpload(downloadedInputType);

            if (Convert.ToString(downloadedInputTypeId) != _valueInputType)
            {
                displayAlerts(clsInformation.misMatchInputTypes, 1);

                return false;
            }

            return true;
        }

        private static int getInputTypeIdForUpload(string downloadedInputType)
        {
            ExcelTool.Worksheet referenceSheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[clsInformation.referenceDataSheet]);

            if (referenceSheet != null)
            {
                int startRowIndex = 2;
                while ((Convert.ToString(referenceSheet.Cells[startRowIndex, "D"].Value)) != "" && (Convert.ToString(referenceSheet.Cells[startRowIndex, "D"].Value)) != null)
                {
                    if ((Convert.ToString(referenceSheet.Cells[startRowIndex, "D"].Value)) == downloadedInputType)
                    {
                        return (Convert.ToInt32(referenceSheet.Cells[startRowIndex, "C"].Value));
                    }

                    startRowIndex++;
                }
            }

            return 0;
        }

        #endregion

        #region checkingValidationsForUpload
        /// <summary>
        /// Checks whether the data meets the validations or not
        /// </summary>
        /// <returns>true or false</returns>
        private static bool checkingValidationsForUpload()
        {

            ExcelTool.Worksheet inputTemplateSheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[clsInformation.productRevenue]);

            inputTemplateSheet.Select();


            if (_dsDownloadData.Tables.Count > 3)
            {
                if (_startRange != null && _endRange != null)
                {
                    _checkRangeForUpload = inputTemplateSheet.Range[_startRange, _endRange] as Excel.Range;

                    #region check Incorrect Formula While Upload
                    try
                    {
                        if (!checkIncorrectFormulaWhileUpload())
                        {
                            return false;
                        }

                    }

                    catch { }
                    #endregion

                    #region check Unwanted Text While Upload
                    try
                    {
                        if (!checkUnwantedTextWhileUpload())
                        {
                            return false;
                        }
                    }
                    catch { }
                    #endregion

                    #region check Data WithIn Specified Range While Upload
                    try
                    {
                        if (!checkDataInSpecifiedRangeWhileUpload())
                        {
                            return false;
                        }
                    }
                    catch { }
                    #endregion

                    #region Check All Cells Required Condition
                    try
                    {
                        if (!checkAllCellsRequiredCondition())
                        {
                            return false;
                        }
                    }
                    catch { }
                    #endregion

                    #region Changing the Number Format
                    // UnProtecting Sheet
                    convertFromSheetNameToProtectAndUnProtect(2, inputTemplateSheet.Name);
                    // For changing the Number format from any other Type to Decimal or Percentages
                    if (_dataTypeValue == "Decimal")
                    {
                        inputTemplateSheet.Range[_startRange, _endRange].NumberFormat = "#,##0.00";
                        inputTemplateSheet.Range[_startRange, _endRange].Interior.ColorIndex = 19;
                        inputTemplateSheet.Range[_startRange, _endRange].Font.Bold = false;
                        inputTemplateSheet.Range[_startRange, _endRange].Font.Color = ColorTranslator.FromHtml("#000");

                    }
                    else if (_dataTypeValue == "Percent")
                    {
                        inputTemplateSheet.Range[_startRange, _endRange].NumberFormat = "##0.00%";
                        inputTemplateSheet.Range[_startRange, _endRange].Interior.ColorIndex = 19;
                        inputTemplateSheet.Range[_startRange, _endRange].Font.Bold = false;
                        inputTemplateSheet.Range[_startRange, _endRange].Font.Color = ColorTranslator.FromHtml("#000");


                    }

                    // aligining the text to right in cells.
                    inputTemplateSheet.Range[_startRange, _endRange].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                    #endregion

                    #region check Blank Cells Condition
                    try
                    {
                        // For Protecting the Sheet
                        convertFromSheetNameToProtectAndUnProtect(1, inputTemplateSheet.Name);

                        if (!checkBlankCellsCondition())
                        {
                            return false;
                        }

                        return true;

                    }
                    catch { }
                    finally
                    {
                        //if (_txtProcess == clsInformation.tcpu)
                        //{
                        //    ExcelTool.Worksheet sheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[_productRevenue.Name]);
                        //    string[] columns = Convert.ToString(FAST._dsDownloadData.Tables[1].Rows[0].ItemArray[0]).Split(',');
                        //    string ColumnName = clsManageSheet.getColumnName(columns.Length + 1);
                        //    clsManageSheet.lockAndSetLifeTimeValueColumnForMSRP(sheet, clsManageSheet.bodyRowStartingNumber + 1, clsManageSheet.bodyRowStartingNumber + _dsDownloadData.Tables[2].Rows.Count, ColumnName);
                        //}  
                        // For Protecting the Sheet
                        convertFromSheetNameToProtectAndUnProtect(1, inputTemplateSheet.Name);
                    }
                    #endregion

                    return true;
                }
                else
                {
                    displayAlerts(clsInformation.scenarioClose, 1);
                    return false;
                }
            }
            else
            {
                displayAlerts("no data found to upload", 1);
                return false;
            }


        }
        #endregion

        #region OperationsforUpload
        /// <summary>
        /// Checks whether any incomplete formulas are specified for the cells
        /// </summary>
        /// <returns>true or false</returns>

        private static bool checkIncorrectFormulaWhileUpload()
        {
            #region Checking if wrong formula or text and date values are available or not


            try
            {
                // UnProtecting
                convertFromSheetNameToProtectAndUnProtect(2, _productRevenue.Name);

                Excel.Range rbn = _checkRangeForUpload.SpecialCells(Excel.XlCellType.xlCellTypeFormulas, Excel.XlSpecialCellsValue.xlErrors) as Excel.Range;

                rbn.Cells[1, 1].select();

                int unwantedTextCount = rbn.Count;


                if (unwantedTextCount >= 1)
                {

                    rbn.Font.Color = ColorTranslator.FromHtml("#FF0000");
                    rbn.Interior.ColorIndex = 19;
                    rbn.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                    if (_productRevenue.Range[_startRange, _endRange].FormatConditions.Count >= 1)
                    {
                        _productRevenue.Range[_startRange, _endRange].FormatConditions.Delete();
                    }

                    Excel.FormatCondition changeFontColorBlack = (Excel.FormatCondition)(_productRevenue.get_Range(_startRange + ":" + _endRange,
                    Type.Missing).FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlBetween,
                    _maxValue, _minValue, Type.Missing, Type.Missing, Type.Missing, Type.Missing));

                    changeFontColorBlack.Font.Color = ColorTranslator.FromHtml("#000");




                    // Protecting Sheet
                    convertFromSheetNameToProtectAndUnProtect(1, _productRevenue.Name);

                    displayAlerts(clsInformation.textNotAllowed, 1);


                    return false;
                }
            }
            finally
            {


                // Protecting Sheet
                convertFromSheetNameToProtectAndUnProtect(1, _productRevenue.Name);
            }


            return true;



            #endregion
        }

        /// <summary>
        /// Checking whether there is any unwanted text or not
        /// </summary>
        /// <returns>true or false</returns>


        private static bool checkUnwantedTextWhileUpload()
        {
            _checkRangeForUpload = _productRevenue.Range[_startRange, _endRange] as Excel.Range;

            try
            {
                if (_startRange != _endRange)
                {
                    // Un-Protecting Sheet
                    convertFromSheetNameToProtectAndUnProtect(2, _productRevenue.Name);

                    Excel.Range rbn = _checkRangeForUpload.SpecialCells(Excel.XlCellType.xlCellTypeConstants, Excel.XlSpecialCellsValue.xlTextValues) as Excel.Range;
                    rbn.Cells[1, 1].select();
                    int unwantedTextCount = rbn.Count;


                    if (unwantedTextCount >= 1)
                    {


                        rbn.Font.Color = ColorTranslator.FromHtml("#FF0000"); //#FF0000
                        rbn.Interior.ColorIndex = 19;
                        rbn.Font.Bold = false;
                        rbn.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;



                        if (_dataTypeValue == "Decimal")
                        {
                            rbn.NumberFormat = "#,##0.00";

                            Excel.FormatCondition changeFontColorBlack = (Excel.FormatCondition)(_productRevenue.get_Range(_startRange + ":" + _endRange,
                            Type.Missing).FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlBetween,
                            _maxValue, _minValue, Type.Missing, Type.Missing, Type.Missing, Type.Missing));

                            changeFontColorBlack.Font.Color = ColorTranslator.FromHtml("#000");
                            changeFontColorBlack.Interior.ColorIndex = 19;
                        }
                        else if (_dataTypeValue == "Percent")
                        {
                            rbn.NumberFormat = "##0.00%";


                            if (_productRevenue.Range[_startRange, _endRange].FormatConditions.Count >= 1)
                            {
                                _productRevenue.Range[_startRange, _endRange].FormatConditions.Delete();
                            }

                            Excel.FormatCondition changeFontColorBlack = (Excel.FormatCondition)(_productRevenue.get_Range(_startRange + ":" + _endRange,
                          Type.Missing).FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlBetween,
                          _maxValue, _minValue, Type.Missing, Type.Missing, Type.Missing, Type.Missing));

                            changeFontColorBlack.Font.Color = ColorTranslator.FromHtml("#000");
                            changeFontColorBlack.Interior.ColorIndex = 19;
                        }

                        if (_txtProcess == clsInformation.tcpuView)
                        {
                            clsManageSheet.lockAndSetLifeTimeValueColumnForMSRP(clsInformation.productRevenue);
                        }


                        // Protecting Sheet
                        convertFromSheetNameToProtectAndUnProtect(1, _productRevenue.Name);


                        displayAlerts(clsInformation.textNotAllowed, 1);
                        return false;
                    }
                }
                else
                {
                    string verifyValue = Convert.ToString(_productRevenue.Range[_startRange, _endRange].Value);

                    if (!IsDigitsOnly(verifyValue))
                    {

                        // Unprotecting Sheet
                        convertFromSheetNameToProtectAndUnProtect(2, _productRevenue.Name);

                        _productRevenue.Cells[_startRange, _endRange].Font.Color = ColorTranslator.FromHtml("#FF0000");

                        Excel.FormatCondition changeFontColorBlack = (Excel.FormatCondition)(_productRevenue.get_Range(_startRange + ":" + _endRange,
                          Type.Missing).FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlBetween,
                          _maxValue, _minValue, Type.Missing, Type.Missing, Type.Missing, Type.Missing));

                        changeFontColorBlack.Font.Color = ColorTranslator.FromHtml("#000");
                        changeFontColorBlack.Interior.ColorIndex = 19;

                        // Protecting Sheet
                        convertFromSheetNameToProtectAndUnProtect(1, _productRevenue.Name);
                        displayAlerts(clsInformation.textNotAllowed, 1);
                        return false;
                    }
                }
            }
            finally
            {

                // Protecting Sheet
                convertFromSheetNameToProtectAndUnProtect(1, _productRevenue.Name);
            }

            return true;
        }

        private static bool IsDigitsOnly(string str)
        {
            if (str != null)
            {
                foreach (char c in str)
                {
                    if ((c < '0' || c > '9') && c != '.')
                        return false;
                }
            }
            return true;
        }

        /// <summary>
        /// Checks whether the edited data is in specified range or not
        /// </summary>
        /// <returns>true or false</returns>

        private static bool checkDataInSpecifiedRangeWhileUpload()
        {
            if (_productRevenue == null)
                clsManageSheet.buildSheet(ref _productRevenue, clsInformation.productRevenue);


            #region condition to check if the data values are greater or lesser than the specified range

            if (_dataTypeValue == clsInformation.decimalType)
            {

                double exceedRange = Globals.ThisAddIn.Application.WorksheetFunction.CountIf(_checkRangeForUpload, ">" + _maxValue);
                double exceedMinRange = Globals.ThisAddIn.Application.WorksheetFunction.CountIf(_checkRangeForUpload, "<" + _minValue);

                if (exceedRange >= 1 || exceedMinRange >= 1)
                {

                    // UnProtecting Sheet
                    convertFromSheetNameToProtectAndUnProtect(2, _productRevenue.Name);

                    // If any previous formatting conditions avaialable deleting them.
                    if (_productRevenue.Range[_startRange, _endRange].FormatConditions.Count >= 1)
                    {
                        _productRevenue.Range[_startRange, _endRange].FormatConditions.Delete();
                    }
                    Excel.FormatCondition changeFontColorBlackforDecimal = (Excel.FormatCondition)(_productRevenue.get_Range(_startRange + ":" + _endRange,
                     Type.Missing).FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlBetween,
                     _maxValue, _minValue, Type.Missing, Type.Missing, Type.Missing, Type.Missing));

                    changeFontColorBlackforDecimal.Font.Color = ColorTranslator.FromHtml("#000");

                    // For checking the maximum value greater than condition
                    Excel.FormatCondition format1 = (Excel.FormatCondition)(_productRevenue.get_Range(_startRange + ":" + _endRange,
                      Type.Missing).FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue,
                       Excel.XlFormatConditionOperator.xlGreater,
                     _maxValue, Type.Missing, Type.Missing, Type.Missing, Type.Missing));

                    // For checking the maximum value greater than condition
                    Excel.FormatCondition format2 = (Excel.FormatCondition)(_productRevenue.get_Range(_startRange + ":" + _endRange,
                      Type.Missing).FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue,
                       Excel.XlFormatConditionOperator.xlLess,
                     _minValue, Type.Missing, Type.Missing, Type.Missing, Type.Missing));
                    

                    if (_valueInputType != "0" && _txtProcess != clsInformation.tcpuView)
                    {
                        if (Convert.ToString(_minValue).Contains(".") && Convert.ToString(_maxValue).Contains("."))
                        {
                            displayAlerts(clsInformation.cellValueExceeded + (_minValue) + " and " + _maxValue, 2);
                        }
                        else if (Convert.ToString(_minValue) != "" && Convert.ToString(_maxValue).Contains("."))
                        {
                            displayAlerts(clsInformation.cellValueExceeded + (_minValue) + " and " + _maxValue, 2);
                        }
                        else if (Convert.ToString(_minValue).Contains("."))
                        {
                            displayAlerts(clsInformation.cellValueExceeded + (_minValue) + " and " + _maxValue.ToString("#,##0"), 2);
                        }
                        else
                        {
                            displayAlerts(clsInformation.cellValueExceeded + (_minValue).ToString("#,#00") + " and " + _maxValue.ToString("#,##0"), 2);
                        }

                        format1.Font.Color = ColorTranslator.FromHtml("#FF0000");   // For Red #ffd700
                        format2.Font.Color = ColorTranslator.FromHtml("#FF0000");   // For Red #ffd700


                        checkForRangeExceedValue(clsInformation.decimalType, exceedRange, exceedMinRange);

                        if (_txtProcess == clsInformation.tcpuView)
                        {
                            clsManageSheet.lockAndSetLifeTimeValueColumnForMSRP(clsInformation.productRevenue);
                        }

                        // Protecting Sheet
                        convertFromSheetNameToProtectAndUnProtect(1, _productRevenue.Name);
                        return false;
                    }

                    if (_valueInputType == "0" && _txtProcess == clsInformation.tcpuView)
                    {
                        if (_txtProcess == clsInformation.tcpuView)
                        {
                            clsManageSheet.lockAndSetLifeTimeValueColumnForMSRP(clsInformation.productRevenue);
                        }

                        // Protecting Sheet
                        convertFromSheetNameToProtectAndUnProtect(1, _productRevenue.Name);
                        return true;
                    }

                    //if (Convert.ToString(_minValue).Contains(".") && Convert.ToString(_maxValue).Contains("."))
                    //{
                    //    displayAlerts(clsInformation.cellValueExceeded + (_minValue) + " and " + _maxValue, 2);
                    //}
                    //else if (Convert.ToString(_minValue) != "" && Convert.ToString(_maxValue).Contains("."))
                    //{
                    //    displayAlerts(clsInformation.cellValueExceeded + (_minValue) + " and " + _maxValue, 2);
                    //}
                    //else if (Convert.ToString(_minValue).Contains("."))
                    //{
                    //    displayAlerts(clsInformation.cellValueExceeded + (_minValue) + " and " + _maxValue.ToString("#,##0"), 2);
                    //}
                    //else
                    //{
                    //    displayAlerts(clsInformation.cellValueExceeded + (_minValue).ToString("#,#00") + " and " + _maxValue.ToString("#,##0"), 2);
                    //}

                    //format1.Font.Color = ColorTranslator.FromHtml("#FF0000");   // For Red #ffd700
                    //format2.Font.Color = ColorTranslator.FromHtml("#FF0000");   // For Red #ffd700

                    // For Selecting the First Cell in the Exceeded Range
                    //checkForRangeExceedValue(clsInformation.decimalType, exceedRange, exceedMinRange);

                    //if (_txtProcess == clsInformation.tcpuView)
                    //{
                    //    clsManageSheet.lockAndSetLifeTimeValueColumnForMSRP(clsInformation.productRevenue);
                    //}

                    //// Protecting Sheet
                    //convertFromSheetNameToProtectAndUnProtect(1, _productRevenue.Name);
                    //return false;
                }
            }
            //else if (FAST._dataTypeValue == clsInformation.percentType && _valueInputType != "0")
            else if (FAST._dataTypeValue == clsInformation.percentType)
            {

                double exceedRange = Globals.ThisAddIn.Application.WorksheetFunction.CountIf(_checkRangeForUpload, ">" + Math.Round(_maxValue * 100) + "%");
                double exceedMinRange = Globals.ThisAddIn.Application.WorksheetFunction.CountIf(_checkRangeForUpload, "<" + Math.Round(_minValue * 100) + "%");

                if (exceedRange >= 1 || exceedMinRange >= 1)
                {
                    // Un-Protecting Sheet
                    convertFromSheetNameToProtectAndUnProtect(2, _productRevenue.Name);

                    // If any previous formatting conditions avaialable deleting them.
                    if (_productRevenue.Range[_startRange, _endRange].FormatConditions.Count >= 1)
                    {
                        _productRevenue.Range[_startRange, _endRange].FormatConditions.Delete();
                    }

                    Excel.FormatCondition changeFontColorBlackforDecimal = (Excel.FormatCondition)(_productRevenue.get_Range(_startRange + ":" + _endRange,
                   Type.Missing).FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlBetween,
                   _maxValue, _minValue, Type.Missing, Type.Missing, Type.Missing, Type.Missing));

                    changeFontColorBlackforDecimal.Font.Color = ColorTranslator.FromHtml("#000");

                    // For checking the maximum value greater than condition
                    Excel.FormatCondition format1 = (Excel.FormatCondition)(_productRevenue.get_Range(_startRange + ":" + _endRange,
                      Type.Missing).FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue,
                      Excel.XlFormatConditionOperator.xlGreater,
                     Convert.ToDecimal(Math.Round(_maxValue * 100)) + "%", Type.Missing, Type.Missing, Type.Missing, Type.Missing));

                    Excel.FormatCondition format2 = (Excel.FormatCondition)(_productRevenue.get_Range(_startRange + ":" + _endRange,
                     Type.Missing).FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue,
                      Excel.XlFormatConditionOperator.xlLess,
                    Convert.ToDecimal(Math.Round(_minValue * 100)) + "%", Type.Missing, Type.Missing, Type.Missing, Type.Missing));

                    displayAlerts(clsInformation.cellRange + Math.Round(_minValue * 100) + "% and " + Math.Round(_maxValue * 100) + "%", 2);
                    format1.Font.Color = ColorTranslator.FromHtml("#FF0000");   // For Red #ffd700
                    format2.Font.Color = ColorTranslator.FromHtml("#FF0000");


                    // For Selecting the First cell in the Exceeded Range
                    checkForRangeExceedValue(clsInformation.percentType, exceedRange, exceedMinRange);

                    // Protecting Sheet
                    convertFromSheetNameToProtectAndUnProtect(1, _productRevenue.Name);
                    return false;
                }
            }

            #endregion
            return true;
        }


        /// <summary>
        /// Checks all cells are required or not
        /// </summary>
        /// <returns>true or false</returns>

        private static bool checkAllCellsRequiredCondition()
        {
            #region checking all cells required condition

            // For blankcells checking condition
            double rngcnt = Globals.ThisAddIn.Application.WorksheetFunction.CountBlank(_checkRangeForUpload);

            if (_allFieldsRequiredConditionForUpload == 0)
            {
                return true;
            }
            else
            {
                if (Convert.ToInt32(rngcnt) >= 1)
                {
                    displayAlerts(clsInformation.cellBlankValues, 2);

                    // UnProtecting Sheet
                    convertFromSheetNameToProtectAndUnProtect(2, _productRevenue.Name);

                    // If any previous formatting conditions avaialable deleting them.
                    if (_productRevenue.Range[_startRange, _endRange].FormatConditions.Count >= 1)
                    {
                        _productRevenue.Range[_startRange, _endRange].FormatConditions.Delete();
                    }

                    Excel.FormatCondition format = (Excel.FormatCondition)(_productRevenue.get_Range(_startRange + ":" + _endRange,
                    Type.Missing).FormatConditions.Add(Excel.XlFormatConditionType.xlBlanksCondition, Excel.XlFormatConditionOperator.xlEqual,
                    0, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing));

                    format.Interior.Color = null;
                    format.Font.Bold = false;

                    format.Interior.Color = ColorTranslator.FromHtml("#ffd700");   // For Gold #ffd700
                                                                                   //_colorFormatForOffline = true;
                                                                                   //clsDataSheet.updateColorFormatValue();

                    Excel.FormatCondition changeFontColorBlackforDecimal = (Excel.FormatCondition)(_productRevenue.get_Range(_startRange + ":" + _endRange,
                     Type.Missing).FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlBetween,
                     _maxValue, _minValue, Type.Missing, Type.Missing, Type.Missing, Type.Missing));

                    changeFontColorBlackforDecimal.Font.Color = ColorTranslator.FromHtml("#000");

                    if (_txtProcess == clsInformation.tcpuView)
                    {
                        clsManageSheet.lockAndSetLifeTimeValueColumnForMSRP(clsInformation.productRevenue);
                    }

                    // Protecting Sheet
                    convertFromSheetNameToProtectAndUnProtect(1, _productRevenue.Name);

                }
                else
                {
                    //_colorFormatForOffline = false;
                    //clsDataSheet.updateColorFormatValue();
                    return true;

                }

                return false;
            }

            #endregion
        }

        /// <summary>
        /// Checks all cells and if any blank cells available, provides alert for the User
        /// </summary>
        /// <returns>true or false</returns>

        private static bool checkBlankCellsCondition()
        {
            // For blankcells checking condition
            double rngcnt = Globals.ThisAddIn.Application.WorksheetFunction.CountBlank(_checkRangeForUpload);

            #region
            if (Convert.ToInt32(rngcnt) >= 1)
            {
                // Null Values are Available Here, so showing the dialog box here
                DialogResult result1 = MessageBox.Show(clsInformation.cellsNoValue, "Amazon Devices - Finance Input Tool", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result1 == DialogResult.Yes)
                {
                    //code for Yes
                    return true;
                }
                else if (result1 == DialogResult.No)
                {
                    //code for No

                    if (_productRevenue == null)
                        clsManageSheet.buildSheet(ref _productRevenue, clsInformation.productRevenue);

                    // Un-Protecting Sheet
                    convertFromSheetNameToProtectAndUnProtect(2, _productRevenue.Name);

                    // If any previous formatting conditions avaialable deleting them.
                    if (_productRevenue.Range[_startRange, _endRange].FormatConditions.Count >= 1)
                    {
                        _productRevenue.Range[_startRange, _endRange].FormatConditions.Delete();
                    }

                    Excel.FormatCondition format = (Excel.FormatCondition)(_productRevenue.get_Range(_startRange + ":" + _endRange,
                    Type.Missing).FormatConditions.Add(Excel.XlFormatConditionType.xlBlanksCondition, Excel.XlFormatConditionOperator.xlEqual,
                    0, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing));

                    if (_refSheetData == null)
                        _refSheetData = (Excel.Worksheet)Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[clsInformation.referenceDataSheet]);

                    //_colorFormatForOffline = true;


                    //clsDataSheet.updateColorFormatValue();
                    format.Interior.Color = ColorTranslator.FromHtml("#ffd700");   // For Gold #ffd700

                    Excel.FormatCondition changeFontColorBlack = (Excel.FormatCondition)(_productRevenue.get_Range(_startRange + ":" + _endRange,
                 Type.Missing).FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlBetween,
                 _maxValue + "%", _minValue + "%", Type.Missing, Type.Missing, Type.Missing, Type.Missing));

                    changeFontColorBlack.Font.Color = ColorTranslator.FromHtml("#000");
                    changeFontColorBlack.Font.Bold = false;
                    changeFontColorBlack.Interior.ColorIndex = 19;

                    Excel.FormatCondition changeFontColorBlackforDecimal = (Excel.FormatCondition)(_productRevenue.get_Range(_startRange + ":" + _endRange,
                 Type.Missing).FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlBetween,
                 _maxValue, _minValue, Type.Missing, Type.Missing, Type.Missing, Type.Missing));

                    changeFontColorBlackforDecimal.Font.Color = ColorTranslator.FromHtml("#000");
                    changeFontColorBlack.Interior.ColorIndex = 19;
                    changeFontColorBlack.Font.Bold = false;


                    if (_txtProcess == clsInformation.tcpuView)
                    {
                        clsManageSheet.lockAndSetLifeTimeValueColumnForMSRP(clsInformation.productRevenue);
                    }


                    // Protecting Sheet
                    convertFromSheetNameToProtectAndUnProtect(1, _productRevenue.Name);

                    return false;
                }
            }
            else
            {
                return true;
            }
            #endregion

            return true;
        }

        /// <summary>
        /// Used for Removing the Blank cells background color
        /// </summary>

        private void removeBackGroundColorForBalnkCells()
        {
            try
            {
                updateEvents(false);
                // Protecting Sheet
                convertFromSheetNameToProtectAndUnProtect(2, _productRevenue.Name);

                if (_productRevenue.Range[_startRange, _endRange].FormatConditions.Count >= 1)
                {
                    _productRevenue.Range[_startRange, _endRange].FormatConditions.Delete();
                }

                Excel.FormatCondition format = (Excel.FormatCondition)(_productRevenue.get_Range(_startRange + ":" + _endRange,
                  Type.Missing).FormatConditions.Add(Excel.XlFormatConditionType.xlBlanksCondition, Excel.XlFormatConditionOperator.xlEqual,
                  0, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing));

                format.Interior.Color = null;
                format.Font.Bold = false;
                format.Interior.ColorIndex = 19;

                if (_txtProcess == clsInformation.tcpuView)
                {
                    clsManageSheet.lockAndSetLifeTimeValueColumnForMSRP(clsInformation.productRevenue);
                }


            }
            catch { }
            finally
            {
                // Protecting Sheet
                convertFromSheetNameToProtectAndUnProtect(1, _productRevenue.Name);

                updateEvents(true);
            }

        }

        #endregion

        #region Converting sheetName and Protecting and UnProtecting
        private static void convertFromSheetNameToProtectAndUnProtect(byte type, string sheetName)
        {
            ExcelTool.Worksheet sheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[sheetName]);
            if (type == 1)
            {
                clsManageSheet.protectSheet(sheet);
            }
            else if (type == 2)
            {
                clsManageSheet.unProtect(sheet);
            }
        }
        #endregion

        #region Select First Cell in the range

        #region 2nd Approach

        public static void checkForRangeExceedValue(string type, double exceedMaxRange, double exceedMinRange)
        {
            if (IsUninstall())
            {
                return;
            }

            var startRange = Regex.Replace(_startRange, @"[\d-]", string.Empty);

            var endRange = Regex.Replace(_endRange, @"[\d-]", string.Empty);


            int startValue = _txtProcess == "Accounting View" ? 11 : 12;
            byte endValue = Convert.ToByte(GetColumnNumber(endRange));


            switch (type)
            {
                case "Decimal":

                    // For Minimum Value
                    if (exceedMinRange >= 1)
                    {
                        selectFirstCell(startRange, endRange, startValue, endValue, false);
                        return;
                    }


                    // For Maximum Value
                    if (exceedMaxRange >= 1)
                    {
                        selectFirstCell(startRange, endRange, startValue, endValue, true);

                        return;
                    }

                    break;

                case "Percent":


                    // For Minimum Value
                    if (exceedMinRange >= 1)
                    {
                        selectFirstCell(startRange, endRange, startValue, endValue, false);

                        return;
                    }


                    // For Maximum Value
                    if (exceedMaxRange >= 1)
                    {
                        selectFirstCell(startRange, endRange, startValue, endValue, true);

                        return;
                    }
                    break;
            }


        }

        private static void selectFirstCell(string startRange, string endRange, int startValue, byte endValue, bool value)
        {
            double rowIndex = 0, colIndex;

            while (startRange != endRange || startRange == endRange)
            {
                string getColumnName = clsManageSheet.getColumnName(startValue);
                if (startValue <= endValue)
                {
                    double rangeValue = 0;

                    // For Maximum Value
                    if (value)
                    {
                        rangeValue = Globals.ThisAddIn.Application.WorksheetFunction.Max(_productRevenue.get_Range(getColumnName + Convert.ToString(clsManageSheet.bodyRowStartingNumber + 1), getColumnName + (clsManageSheet.bodyRowStartingNumber + _dataSourceLength)));

                        if (rangeValue > Convert.ToDouble(_maxValue))
                        {
                            Excel.Range currentRange = null;

                            // For Decimal
                            if (_dataTypeValue == clsInformation.decimalType)
                            {
                                currentRange = findRange(rangeValue);
                            }
                            // For Percent Type
                            else if (_dataTypeValue == clsInformation.percentType)
                            {
                                currentRange = findRange(rangeValue * 100);
                            }


                            if (currentRange != null)
                            {
                                rowIndex = currentRange.Row; colIndex = currentRange.Column;

                                while (rowIndex != clsManageSheet.bodyRowStartingNumber)
                                {

                                    if (Convert.ToDouble(_productRevenue.Cells[rowIndex, colIndex].Value) > Convert.ToDouble(_maxValue))
                                        _productRevenue.Cells[rowIndex, colIndex].Select();

                                    rowIndex--;

                                    if (rowIndex == clsManageSheet.bodyRowStartingNumber)
                                    {
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    // For Minumum Value
                    else if (!value)
                    {
                        rangeValue = Globals.ThisAddIn.Application.WorksheetFunction.Min(_productRevenue.get_Range(getColumnName + Convert.ToString(clsManageSheet.bodyRowStartingNumber + 1), getColumnName + (clsManageSheet.bodyRowStartingNumber + _dataSourceLength)));


                        if (rangeValue < Convert.ToDouble(_minValue))
                        {
                            Excel.Range currentRange = null;

                            // For Decimal
                            if (_dataTypeValue == clsInformation.decimalType)
                            {
                                currentRange = findRange(rangeValue);
                            }
                            // For Percent Type
                            else if (_dataTypeValue == clsInformation.percentType)
                            {
                                currentRange = findRange(rangeValue * 100);
                            }


                            if (currentRange != null)
                            {
                                rowIndex = currentRange.Row; colIndex = currentRange.Column;

                                while (rowIndex != clsManageSheet.bodyRowStartingNumber + 1)
                                {

                                    if (Convert.ToDouble(_productRevenue.Cells[rowIndex, colIndex].Value) < Convert.ToDouble(_minValue))
                                        _productRevenue.Cells[rowIndex, colIndex].Select();

                                    rowIndex--;

                                    if (rowIndex == clsManageSheet.bodyRowStartingNumber + 1)
                                    {
                                        break;
                                    }
                                }
                            }
                        }
                    }

                }
                else
                {
                    break;
                }

                if (rowIndex == clsManageSheet.bodyRowStartingNumber + 1 && rowIndex != 0)
                {
                    break;
                }

                startValue++;

                startRange = clsManageSheet.getColumnName(startValue);
            }



        }

        public static Excel.Range findRange(double findValue)
        {
            Excel.Range identifiedRange = null;

            identifiedRange = _checkRangeForUpload.Find(findValue);

            return identifiedRange;
        }

        public static int GetColumnNumber(string name)
        {
            int number = 0;
            int pow = 1;
            for (int i = name.Length - 1; i >= 0; i--)
            {
                number += (name[i] - 'A' + 1) * pow;
                pow *= 26;
            }

            return number;
        }
        #endregion


        #endregion

        #region uninstall code

        public static Boolean IsUninstall()
        {

            Microsoft.Win32.RegistryKey Key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Office\Excel\Addins\FAST");
            if (Key != null)
            {


                return false;
            }
            else
            {
                timer.Enabled = false;
                timer.Stop();
                MessageBox.Show("The add-in has been uninstalled, Excel will be closed in 3 secs");
                afterUninstalltimer.Interval = 1000;
                afterUninstalltimer.Enabled = true;
                afterUninstalltimer.Tick += new System.EventHandler(UninstallTimer_Tick);
                return true;
            }
        }

        private static void UninstallTimer_Tick(object sender, EventArgs e)
        {
            afterUninstalltimer.Stop();
            MessageBox.Show("Excel Will be closed");
            Globals.ThisAddIn.Application.Quit();

        }
        #endregion

        #region Update Refresh Control
        public static void updateControl()
        {

            if (Globals.ThisAddIn.Application.ActiveWorkbook == null)
                return;
                

            ExcelTool.Workbook wrkbk = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook);
           
            string promoReferenceSheet = null;
            foreach (Excel.Worksheet sheet in wrkbk.Sheets)
            {
                if (sheet.Name == clsInformation.referencePromo)
                    promoReferenceSheet = sheet.Name;

                //MessageBox.Show(sheet.Name);
            }

            if (promoReferenceSheet != null)
            {
                if (isDownloadEnabled)
                {
                    ClsPromotions.promotionsOnOpen();
                }
                

                //added by anwesh 08/22/2017

                Worksheet promoOfflineReferenceSheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[clsInformation.referencePromo]);

                //string download = Convert.ToString(promoOfflineReferenceSheet.Cells[2, 21].Value);
                
                foreach (RibbonGroup group in Globals.Ribbons.Ribbon1.tab1.Groups)
                {

                    if (group.Name == "grpOperations")
                    {
                        foreach (RibbonControl control in group.Items)
                        {   
                            //anwesh 08/22/2019
                            if (control.Name == "btnDownloadData")
                            {
                                RibbonButton btnControl = (RibbonButton)control;
                                btnControl.Enabled = false;
                                // btnControl.Enabled = Convert.ToBoolean(promoOfflineReferenceSheet.Cells[2, 20].Value);
                            }
                            if (control.Name == "btnUploadData")//anwesh 08/22/2019
                            {
                                RibbonButton btnControl = (RibbonButton)control;
                                btnControl.Enabled = Convert.ToBoolean(promoOfflineReferenceSheet.Cells[2, 21].Value);
                            }
                            if (control.Name == "btnRefresh")
                            {
                                RibbonButton btnControl = (RibbonButton)control;
                                btnControl.Label = "Refresh VDP/TCPU Data";
                                btnControl.Enabled = Convert.ToBoolean(promoOfflineReferenceSheet.Cells[2, 22].Value);
                            }
                            if (control.Name == "btnRefreshBransonData")
                            {
                                RibbonButton btnControl = (RibbonButton)control;
                                btnControl.Visible = true;
                                btnControl.Enabled = Convert.ToBoolean(promoOfflineReferenceSheet.Cells[2, 23].Value);
                            }
                            if (control.Name == "btnContactSupport")
                            {
                                RibbonButton btnControl = (RibbonButton)control;
                                btnControl.Visible = true;
                                btnControl.Enabled = true;
                            }
                        }
                    }

                    if (group.Name == "grpReports")
                    {
                        group.Visible = false;
                    }
                }
            }

            }

        #endregion

        #region Test Sheet

        private void testSheet()
        {
            //DataTable dt = new DataTable();
            //dt.Columns.Add("RUN_DATE");
            //dt.Columns.Add("PROMOTION_TYPE");


            //dt.Rows.Add("2017-12-25 07:21:09","TPR");
            //dt.Rows.Add("2017-12-25 07:21:09","TPR");
            //dt.Rows.Add("MONTOYA", "TPR");

            //clsManageSheet.buildSheet(ref _bransonPromotions, clsInformation.bransonPromotions);



            //Worksheet sheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[clsInformation.bransonPromotions]);

            ////Globals.ThisAddIn.Application.AutoCorrect.AutoFillFormulasInLists = false;

            //Excel.Range listObjectControlRange = sheet.get_Range("B10", "B13") as Excel.Range;

            //string listObjectName = "List-Object-" + sheet.Name + clsInformation.bransonPromotions;
            //ExcelTool.ListObject lo = sheet.Controls.AddListObject(listObjectControlRange, listObjectName);

            //string[] columnNames = { "RUN_DATE", "PROMOTION_TYPE"};

            //BindingSource bs = new BindingSource();
            //bs.DataSource = dt;
            //lo.SetDataBinding(bs, "", columnNames);

            //Excel.Range shtTitleRange1 = sheet.Range["B11","B20"] as Excel.Range;

            //shtTitleRange1.NumberFormat = ("yyyy-MM-dd h:mm:ss");

            //int length = 10;

            //sheet.Cells["11", "B"].ColumnWidth = length + 12;
        }
        #endregion
    }
}


