namespace Test_WorkBookOpen.Classes
{
    class clsInformation
    {

        #region Global Decleration

        public const string constSheetTitleStartRange = "B2", constSheetTitleEndRange = "C3", promotions = "Promotion Planning", promoConstSheetTitleStartRange = "B2", promoConstSheetTitleEndRange = "C2";

        public const string constSheetSummaryStartRange = "B4", constSheetSummaryEndRange = "C4", promoconstSheetSummaryStartRange = "B3", promoConstSheetSummaryEndRange = "C3";
        public const string headerStartRange = "B10", headerEndRange = "AO10";

        public const string colorBlack = "#000", colorOrange = "#F79A19";

        public static int sheetTitleContinueRowNumber = 5, sheetTitleContinueColumnNumber = 2;

        public const string sheetTitle = "Finance Advanced Smart Technology", filterSummary = "Filter Summary";

        public const string referenceDataSheet = "RefDataSheet", referencePromo = "referencePromo";

        public const string productRevenue = "Input Template";

        public const string productsRevenuePivot = "Input Template Pivot";

        public const string productsAuditReportPivot = "Audit Report Pivot";


        public const string productsRevenueReport = "Variance Report", statistics = "Statistics";

        public const string productsAuditReport = "Audit Report", isReadOnly = "isReadOnly";

        public const string auditReportSheetNull = "Audit Report Sheet is not available to complete the Process";

        public const string statisticsScenarioLabel = "Scenario - Uploads/Downloads", statisticsScenarioTag = "1",
                            statisticsInputTypeLabel = "InputType - Uploads/Downloads", statisticsInputTypeTag = "2",
                            statisticsUserLabel = "User - Uploads/Downloads", statisticsUserTag = "3",
                            statisticsAlerts = "Statistics Created Successfully";

        public const string msrp = "MSRP (Local Currency)", msrpUSD = "MSRP (USD)", countryUs = "US", country = "Country";

        public const string accountingView = "Accounting View", tcpuView = "TCPU View", admin = "Admin", promotionsView = "Promotion Planning";

        public static string[] headersForReferenceSheet = new string[] {"ID", "Value", "Type", "AllFieldsRequired",
                                                            "Minimum Value", "Maximum Value", "InputTypeDescription", "InputDataType", "VarianceFlagType", "VariancePercentage" };

        public static string[] headernames = new string[] {"ProcessId", "ScenarioId", "InputTypeId", "CurrencyConditionId", "IntervalId","ProductLineId",
                                                            "DownloadedProcessValueforOffline", "DownloadScenarioValueforOffline",
                                                            "DownloadInputTypeValueforOffline", "DownloadCurrencyValueforOffline","DownloadIntervalValueForOffline","DownloadProductLineValueForOffline",
                                                            "MinValue", "MaxValue","AllFieldsRequiredConditionForUpload","DataTypeValue",
                                                            "OfflineDataFile","StartRange","EndRange", "UpdateRowNumberForOffline", "UploadColumnStartRange", "txtProcessValue" };

        public static string[] statisticsScenarioHeader = new string[] { "Scenario", "Downloads", "Uploads" },
                               statisticsInputTypeHeader = new string[] { "InputType", "Downloads", "Uploads" },
                               statisticsUserHeader = new string[] { "User", "InputType", "Downloads", "Uploads" };

        public const string listObjectName = "List-Object", formulaColumnName = "Formula";

        public const string productUpdate = "ProductUpdate", dataTableSave = "FAST_",
                            process = "Process", inputType = "InputType", scenario = "Scenario", interval = "Interval", auditScenario = "AuditScenario", account = "Account",
                            previousScenario = "Previous Scenario", currency = "Currency", productLine = "Product Line", currencyCondition = "CurrencyCondition", user = "User", downloadTimeStamp = "Download TimeStamp",
                            scenarioValidations = "ScenarioValidations", ribbonControlId = "TabAddIns", Discription = "Discription", userRole = "UserRole";

        public const string pleaseSelect = "Please Select ", commaSeperator = ", ", dropdowns = " Dropdowns", andSeperator = " & ";

        public const string defaultScenario = "--Select Scenario--",
                            defaultInputType = "--Select Input Type--",
                            defaultCurrency = "--Select Currency--",
                            defaultInterval = "--Select Interval--",
                            defaultProductLine = "--Select ProductLine--",
                            defaultIntervel = "--Select Intervel--",
                            defaultCountry = "--Select Country--",
                            defaultDeviceType = "--Select DeviceType--";

        public const string sheetStatus = "Offline";

        public const string outlookSubject = "AD3 FAST Support", outlookSubject2 = "AD3 FAST TCPU Support", outlookSubject3 = "AD3 FAST Promotion Planning Support", outlookRecepients = "ad3-fast-support@amazon.com", outlookRecepients2 = "ad3-fast-tcpu-support@amazon.com", outlookRecepients3 = "ad3-fast-promo-support@amazon.com";


        public const string decimalType = "Decimal", percentType = "Percent";

        public const string error = "Error";

        public const string displayMessageTitle = "FAST App";

        public const string errorMessage = "Due to Some Issues, unable to complete the Request. \n Please Try Again.";

        public const string userAccessMessage = "Your not allowed for Further Access.";

        public const string normalNetworkIssue = "Please Check the Internet Connection or Connect to Corp Network";

        public const string corpNetworkIssue = "Due to some Network Issue, the Process cannot be completed";

        public const string sheetValid = "The Sheet is Not Validated For Data Building";

        public const string validatedDataBuilding = "The Sheet is Not Validated For Data Building";

        public const string parameterValidation = "Parameters are Invalid for Sheet Creation";

        public const string buildSheetBodyEmpty = "Worksheet cannot be empty when building sheet body";

        public const string creatingTitle = "Sheet object can not be empty when creating the sheet title";

        public const string toolUpdateSuccess = "Tool Updated Succesfully.Please Restart Excel for changes to effect.";

        public const string toolUpdateFail = "Update failed: Exit Code";

        public const string deploymentDownloadException = "The new version of the application cannot be downloaded at this time. \n\nPlease check your network connection, or try again later. Error:";

        public const string invalidDeploymentException = "Cannot check for a new version of the application. The ClickOnce deployment is corrupt. Please redeploy the application and try again. Error:";

        public const string invalidOperationException = "This application cannot be updated. It is likely not a ClickOnce application. Error:";

        public const string updateApplication = "An update is available. Would you like to update the application now?";

        public const string updateAvailable = "Update Available";

        public const string mandatoryUpdate = "This application has detected a mandatory update from your current";

        public const string appInstallRestartUpdate = "The application will now install the update and restart.";

        public const string appUpgrade = "The application has been upgraded, and will now restart.";

        public const string furtherAccess = "Your not allowed for Further Access.";

        public const string editMode = "Excel Sheet is in Edit Mode, Please complete the operation and Try Again.";

        public const string noActiveWorkbook = "Sorry! There is no Active Workbook available. Please Open the Excel Window.";

        public const string noOpenExcelWorkbook = "There is no open Excel workbook. Please Create a new Empty Workbook and click 'Download Data' button";

        public const string clickInitialize = "Please click on Initialize Workbook to complete the Process.";

        public const string changeUploadSuccess = "Your changes have not been uploaded yet. Click 'OK' if you would like to discard your changes and continue to download template; otherwise click 'Cancel' ";

        public const string downloadSuccess = "Template has been Downloaded successfully!", downloadSuccess1 = " Promotion Template has been Downloaded successfully!", noDataDownload = "No Data to Download";

        public const string multipleWorkbooks = "Unable to Complete the Process. Possiblity reasons can be \n\n 1. Multiple Workbooks were Opened. \n 2. Issues with the current sheet.";

        public const string multipleWorkbooksNewly = "Multiple Excel Workbooks are open. Please close all the Previous workbooks and work with the newly opened Workbook.";

        public const string allDropdowns = "Please Select all Dropdown Values";

        public const string allDropdowns1 = "Please Select all Dropdown Values";

        public const string scenarioInputTypedropdown = "Please Select Scenario and Input Type Dropdown Values";

        public const string scenarioInputinterveldropdown = "Please Select Scenario and Intervel Dropdown Values";

        public const string scenarioCurrencydropdown = "Please Select Scenario and Currency Dropdown Values";

        public const string inputTypeCurrencyDropdown = "Please Select Input Type and Currency Dropdown Values";

        public const string inputTypeIntervelDropdown = "Please Select Input Type and Intervel Dropdown Values";

        public const string intervalScenarioDropdown = "Please Select Interval and Scenario Values";

        public const string intervalCurrencyDropdown = "Please Select Interval and Currency Dropdown Values";

        public const string ProductLineScenarioDropdown = "Please Select ProductLine and Scenario Values"; //Added by Sita 

        public const string ProductLineCurrencyDropdown = "Please Select ProductLine and Currency Dropdown Values";//Added by Sita

        public const string intervalinputTypeDropdown = "Please Select Input Type and Interval Dropdown Values";

        public const string ProductLineinputTypeDropdown = "Please Select Input Type and ProductLine Dropdown Values"; // added by Sita

        public const string scenarioDropdown = "Please Select Scenario";

        public const string inputTypeDropdown = "Please Select Input Type";

        public const string intervalDropdown = "Please Select Interval";

        public const string CountryDropdown = "Please Select Country";

        public const string DeviceTypeDropdown = "Please Select DeviceType";

        public const string CountryDeviceTypeDropdown = "Please Select Country and DeviceType values";



        public const string ProductLineDropdown = "Please Select ProductLine"; //added by Si

        public const string DataValidationCannotSet = "Data validation Cannot be Added for the Range as \n Minimum and Maximum Input Values for a Cell.";

        public const string currencyType = "Please Select Currency";

        public const string refreshPivot = "Please complete the Download Data Process to continue with Refresh Pivot", refreshVdpTcpu = "VDP & TCPU Sheets Refreshed Successfully",
            refreshVdpTcpuCheck = "Please complete the Download Data Process to continue with Refresh VDP/TCPU Data", refreshBransonData = "BransonPromotions Sheet Refreshed Successfully";

        public const string verifyDownloadforUpload = "Please complete the Download Data Process for Uploading.";

        public const string scenarioClose = " Scenario is closed. Please contact support to open it for editing.";

        public const string misMatchTypes = "There is a Mismatch with the download and the upload Types. Please Verify Dropdowns and Filter Summary.",
                            misMatchInputTypes = "There is a Mismatch between the downloaded and Uploaded Input Type.\n\n Please Verify Dropdowns and Filter Summary.";

        public const string textNotAllowed = "Text Values are not allowed.";

        public const string cellValueExceeded = "Cell Values should not exceed ";

        public const string cellRange = "Cell Values should not exceed range between ";

        public const string cellsNoValue = "You have not loaded inputs for every month and program configuration.  Do you still wish to upload your inputs?";

        public const string cellBlankValues = "Please enter the values in the blank cells, to complete the upload.";

        public const string noModificationsDone = "No Modifications done to the Data for Uploading";

        public const string thisOldWorkbook = "Unable to Complete the Process as this an old workbook.";

        public const string reportSuccessfull = "Report has been Generated successfully!";

        public const string noDataReport = "No Data to generate the Report", noDataStatistics = "No Data to generate the Statistics";

        public const string changesUploadSuccess = "have been Uploaded to database successfully";

        public const string uploadFail = "Upload Failed!", uploadSuccess = "UploadSuccess", promotionsUploadSuccess = "Inputs on Promotion Planning has been uploaded into database successfully";

        public const string noChangesInTemplateForUpload = "No changes have been made in the Template for upload";

        public const string dataNotUploaded = "Due to Network Issue, Unable to Complete the Operation. Please Try Again.";

        public const string offlineIssue = "Due to some issues, Data modifications cannot be tracked for this workbook in Offline.";

        public const string contactSupport = "Due to some issues unable to open Outlook. Please Try Again.";

        public const string serviceUnavailable = "Service is Unavailable to Complete the Request. Please Try After Sometime.";

        public const string lessThan = "Less Than";

        public const string greaterThan = "Greater Than";

        public const string both = "Both";

        public const string bColWidth = "22";

        public const string aColWidth = "5";

        public static string listObjName = "";

        public static string updateDownload = "Update downloaded, Please restart Excel.";

        public const string PROMO_INPUT_TOOL = "Promo Input Template";
        public const string TCPU = "TCPU";
        public const string VDP = "VDP";
        public const string bransonPromotions = "BransonPromotions";

         public const string bransonHeader = "RUN_DATE,COUNTRY,DEVICE_TYPE,PROGRAM,PROGRAM_TCPU_MAPPING,CHANNEL,PROMOTION_ID,PROMOTION_STATUS,PROMOTION_TYPE,DESCRIPTION,START_DATE,END_DATE,YEAR,IS_HISTORICAL,DISCOUNT_LOCAL,AMAZON_FUNDING_SPLIT,BASELINE_UNITS,LIFT_FORECAST,INCREMENTAL_UNITS_FORECAST,TOTAL_PROMO_UNITS_FORECAST";

        public const string constStartRange = "B";
        public const string constDateFormat = ("yyyy-MM-dd h:mm:ss");



        public const string PROMOTOOL_DEFAULT_ROW_COLOR = "#fdd49b";
        public const string PROMOTOOL_ALTERNATE_ROW_COLOR = "#fdfdfd";
        public const string PROMOTOOL_EDITCOLUMN_ROW_COLOR = "#ffff99";

        public const string AliasId = "AliasId";

        public const string promoOperations = "Operations", enable = "isEnabled", 
            downloadEnable = "Download", uploadEnable = "Upload", refreshPivotEnable = "RefreshPivot",
            refreshBransonEnable = "refreshBranson", refreshTCPUVDPEnable = "refreshTCPUVDP";

        public const string noPromoInputTemplate = "No Promo Input Template Sheet,Click Download and Upload Sheet";
        #endregion
    }
}
