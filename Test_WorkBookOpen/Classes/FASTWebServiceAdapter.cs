using System;
using System.Configuration;
using System.Data;
using System.IO;
using System.Net;


namespace Test_WorkBookOpen.Classes
{
    public class FASTWebServiceAdapter
    {
       // public static string FAST.baseUrl = Convert.ToString(ConfigurationManager.AppSettings["BaseURL"]);
       // public static string FAST.baseUrl = FAST.baseUrl;

        const string getProcessMenuItems = "initializeWorkbook", getDropdownItems = "initializeParameters", getvarinaceReportScenarios = "getVarianceScenarios",
            getDownloadData = "dynamicDownloadData", sendUploadData = "performanceUpload", getVarianceReport = "generateDynamicVarianceReport", getAuditReport = "getAuditReport",
            sendUploadPromo = "UploadPromoData", refreshVdpTcpu = "refreshVDPTCPUData", refreshBransonData = "refreshRockerData", AutoUpdate = "updateFastVersion";

        public static DataSet getDownloadDataForPromotions(string userName, string processId, string countryId, string deviceTypeId, string _txtProcess)
        {
            var serviceRequestUrl = string.Format("{0}/{1}?AliasId={2}&ProcessId={3}&CountryId={4}&deviceTypeId={5}&View={6}", FAST.baseUrl, getDownloadData, userName, processId, countryId, deviceTypeId, FAST._txtProcess);
            return requestUrl(serviceRequestUrl, null);

        }

        public static DataSet sendUploadDataForPromotions(string userName, string valueProcess, string countryId, string deviceTypeId, string _txtProcess, string uploadData)
        {
            var serviceRequestUrl = string.Format("{0}/{1}?View={2}&ProductLineId={3}&IntervalId={4}&CurrencyConditionId={5}&InputTypeId={6}&ScenarioId={7}&AliasId={8}&ProcessId={9}&CountryId={10}&deviceTypeId={11}",
                 FAST.baseUrl, sendUploadData, FAST._txtProcess, "1", "1", "1", "1", "1", userName, Convert.ToInt32(valueProcess), Convert.ToInt32(countryId), Convert.ToInt32(deviceTypeId));

            return requestUrl(serviceRequestUrl, "Upload", uploadData);
        }

        public static DataSet refreshVdpTcpuForPromoInputTool(string userName, string processId, string countryId, string deviceTypeId, string _txtProcess)
        {
            var serviceRequestUrl = string.Format("{0}/{1}?AliasId={2}&ProcessId={3}&CountryId={4}&deviceTypeId={5}&View={6}", FAST.baseUrl, refreshVdpTcpu, userName, processId, countryId, deviceTypeId, FAST._txtProcess);
            return requestUrl(serviceRequestUrl, null);

        }

        public static DataSet refreshBransonDataForPromoInputTool(string userName, string processId, string countryId, string deviceTypeId, string _txtProcess)
        {
            var serviceRequestUrl = string.Format("{0}/{1}?AliasId={2}&ProcessId={3}&CountryId={4}&deviceTypeId={5}&View={6}", FAST.baseUrl, refreshBransonData, userName, processId, countryId, deviceTypeId, FAST._txtProcess);
            return requestUrl(serviceRequestUrl, "Refresh Branson Data");

        }

        public static DataSet updateFastVersionAutoUpdate(string userName, string fastVersion, DateTime updatedTime)
        {
            var serviceRequestUrl = string.Format("{0}/{1}?AliasId={2}&fastVersion={3} &timeStamp={4}", FAST.baseUrl, AutoUpdate, userName, fastVersion, updatedTime);
            return requestUrl(serviceRequestUrl, "Upload");
        }

        #region Request Urls
        public static DataSet getProcessItems(string userName)
        {
            var serviceRequestUrl = string.Format("{0}/{1}?AliasId={2}", FAST.baseUrl, getProcessMenuItems, userName);
            return requestUrl(serviceRequestUrl, "Initialize");

        }
        public static DataSet getDropdownDataItems(string userName, string ProcessId)
        {
            var serviceRequestUrl = string.Format("{0}/{1}?AliasId={2}&ProcessId={3}&View={4}", FAST.baseUrl, getDropdownItems, userName, ProcessId, FAST._txtProcess);
            return requestUrl(serviceRequestUrl, "Generate");

        }

        public static DataSet getVarianceReportMenuItems(string userName, string valueProcess, string valueScenario, string valueProductLine)
        {
            var serviceRequestUrl = string.Format("{0}/{1}?AliasId={2}&ProcessId={3}&ScenarioId={4}&View={5}&ProductLineId={6}", FAST.baseUrl, getvarinaceReportScenarios, userName, valueProcess, valueScenario, FAST._txtProcess, valueProductLine);
            return requestUrl(serviceRequestUrl, null);
        }

        public static DataSet getDownloadDataForUser(string userName, string valueProcess, string valueScenario, string valueInputType, string valueCurrency, string valueInterval, string valueProductLine)
        {
            var serviceRequestUrl = string.Format("{0}/{1}?AliasId={2}&ProcessId={3}&ScenarioId={4}&InputTypeId={5}&CurrencyConditionId={6}&IntervalId={7}&View={8}&ProductLineId={9}",
                                                        FAST.baseUrl, getDownloadData, userName, valueProcess, valueScenario, valueInputType, valueCurrency, valueInterval, FAST._txtProcess, valueProductLine);
            return requestUrl(serviceRequestUrl, null);
        }

        public static DataSet sendUploadDataForUser(string userName, string valueProcess, string valueScenario, string valueInputType, string valueCurrency, string valueInterval, string valueProductLine)
        {
            var serviceRequestUrl = string.Format("{0}/{1}?AliasId={2}&ProcessId={3}&ScenarioId={4}&InputTypeId={5}&CurrencyConditionId={6}&IntervalId={7}&View={8}&ProductLineId={9}",
                                    FAST.baseUrl, sendUploadData, userName, Convert.ToInt32(valueProcess), Convert.ToInt32(valueScenario), Convert.ToInt32(valueInputType), Convert.ToInt32(valueCurrency), Convert.ToInt32(valueInterval), FAST._txtProcess, valueProductLine);
            return requestUrl(serviceRequestUrl, "Upload");
        }


        /// <summary>
        /// Initializing the variance report instance
        /// </summary>
        /// <param name="userName"></param>
        /// <param name="processId"></param>
        /// <param name="scenarioId"></param>
        /// <param name="inputTypeId"></param>
        /// <param name="currencyId"></param>
        /// <param name="varianceScenarioId"></param>
        /// <returns></returns>
        public DataSet GetVarianceReport(string userName, string processId, string scenarioId, string inputTypeId, string currencyId, string varianceScenarioId, string IntervalId, string valueProductLine)
        {
            var serviceRequestUrl = string.Format("{0}/{1}?AliasId={2}&ProcessId={3}&ScenarioId={4}&InputTypeId={5}&CurrencyConditionId={6}&VarianceScenarioId={7}&IntervalId={8}&View={9}&ProductLineId={10}",
                                                    FAST.baseUrl, getVarianceReport, userName, processId, scenarioId, inputTypeId, currencyId, varianceScenarioId, IntervalId, FAST._txtProcess, valueProductLine);
            return requestUrl(serviceRequestUrl, null);
        }

        public static DataSet getAuditScenarioReport(string userName, string processId, string auditScenarioId)
        {
            var serviceRequestUrl = string.Format("{0}/{1}?AliasId={2}&ProcessId={3}&ScenarioId={4}&View={5}", FAST.baseUrl, getAuditReport, userName, processId, auditScenarioId, FAST._txtProcess);
            return requestUrl(serviceRequestUrl, "Audit");


        }
        #endregion


        #region Service Action

        /// <summary>
        /// Used to send data to the service layer
        /// </summary>
        /// <param name="serviceRequestUrl">url will be passed as parameter</param>
        /// <param name="info">specifies upload/download/etc..</param>
        /// <returns></returns>
        private class WebClient : System.Net.WebClient
        {
            public int Timeout { get; set; }

            protected override WebRequest GetWebRequest(Uri uri)
            {
                WebRequest lWebRequest = base.GetWebRequest(uri);
                WebRequest w = base.GetWebRequest(uri);
                w.Timeout = 20 * 60 * 1000;
                return w;
            }
        }


        public static DataSet requestUrl(string serviceRequestUrl, string info, string uploadData = null)
        {
            FAST.web = new WebClient();
            DataSet localDataSet = new DataSet();
            string strFinalUpload = null;

            using (var lWebClient = new WebClient())
                lWebClient.Timeout = 600 * 60 * 1000;


            if (info != "Upload")
            {
                FAST.response = FAST.web.DownloadString(serviceRequestUrl);
                FAST.response = FAST.response.Replace("null", "");
            }
            else
            {
                if (FAST._txtProcess == clsInformation.accountingView)
                {
                    FAST.response = FAST.web.UploadString(serviceRequestUrl, Convert.ToString(clsproductUpdateXMLManager._data));

                    //added by praveen(CSV Response)
                    //strFinalUpload = Convert.ToString(clsproductUpdateXMLManager.sb);
                    //FAST.response = FAST.web.UploadString(serviceRequestUrl, strFinalUpload);
                }
                else if (FAST._txtProcess == clsInformation.tcpuView)
                {

                    strFinalUpload = Convert.ToString(clsproductUpdateXMLManager.sb);
                    FAST.response = FAST.web.UploadString(serviceRequestUrl, strFinalUpload);
                }
                else if (FAST._txtProcess == clsInformation.promotionsView)
                {

                    FAST.response = FAST.web.UploadString(serviceRequestUrl, uploadData);
                }
                else
                {
                    FAST.response = FAST.web.UploadString(serviceRequestUrl, "upload");
                }

            }

            switch (info)
            {
                case "Audit":
                    FAST.response = "<Audit>" + FAST.response + "</Audit>";
                    using (StringReader stringReader = new StringReader(FAST.response))
                    {
                        localDataSet.ReadXml(stringReader);
                    }
                    break;

                case "Initialize":
                    //TODO: Remove Below hard coded one once service is ready
                    //FAST.response = "<InitializeElements><view><Value>Accounting View</Value><Id>1</Id></view><view><Value>TCPU View</Value><Id>2</Id></view><view><Value>Promotions</Value><Id>3</Id></view></InitializeElements> ";

                    if (!FAST.response.Contains("false"))
                    {
                        using (StringReader stringReader = new StringReader(FAST.response))
                        {
                            localDataSet.ReadXml(stringReader);
                        }
                    }
                    break;

                case "Generate":
                    //TODO: Remove Below hard coded one once service is ready
                    //if (string.Equals(FAST._txtProcess, "Promotion Planning", StringComparison.InvariantCultureIgnoreCase))
                    //	FAST.response = "<InitializeElements><Parameters><UserRole>SuperAdmin</UserRole></Parameters><Parameters><isReadOnly>0</isReadOnly></Parameters><Parameters><Type>Country</Type><Value>US</Value><Id>208</Id></Parameters><Parameters><Type>Device Type</Type><Value>XXX</Value><Id>208</Id></Parameters></InitializeElements>";

                    if (FAST.response != "")
                    {
                        using (StringReader stringReader = new StringReader(FAST.response))
                        {
                            DataSet tds = new DataSet();
                            tds.ReadXml(stringReader);

                            if (tds.Tables.Count > 0 && (tds.Tables[0].Rows.Count >= 2))
                            {
                                if (FAST._txtProcess == clsInformation.tcpuView)
                                {
                                    // rearranging the dataTable Columns
                                    tds.Tables[0].Columns["VarianceValue"].SetOrdinal(tds.Tables[0].Columns.Count - 1);
                                }


                                DataTable scenarioDataTable, currencyConditionDataTable, inputTypeDataTable, intervalDataTable, userRoleDataTable, auditScenarioTable, productLineDataTable, countryDataTable, deviceTypeDataTable, promoOperationsTable;
                                scenarioDataTable = tds.Tables[0].Clone();
                                currencyConditionDataTable = tds.Tables[0].Clone();
                                inputTypeDataTable = tds.Tables[0].Clone();
                                intervalDataTable = tds.Tables[0].Clone();
                                userRoleDataTable = tds.Tables[0].Clone();
                                auditScenarioTable = tds.Tables[0].Clone();
                                productLineDataTable = tds.Tables[0].Clone();
                                countryDataTable = tds.Tables[0].Clone();
                                deviceTypeDataTable = tds.Tables[0].Clone();
                                int userRoleColumnCount = 0;

                                for (int i = 0; i < tds.Tables[0].Columns.Count; i++)
                                {
                                    if (tds.Tables[0].Columns[i].ColumnName == "Type")
                                    {
                                        FAST.labelItemArrayNumber = i;
                                    }
                                    if (tds.Tables[0].Columns[i].ColumnName == "Id")
                                    {
                                        FAST.tagItemArrayNumber = i;
                                    }
                                    if (tds.Tables[0].Columns[i].ColumnName == "UserRole")
                                    {
                                        userRoleColumnCount = i;
                                    }
                                }

                                foreach (DataRow drtableOld in tds.Tables[0].Rows)
                                {
                                    string value = Convert.ToString(drtableOld.ItemArray[FAST.labelItemArrayNumber]);
                                    // Removed an if condtion w.r.t TCPU Code FAST Phase -2 at line number - 247// sainadup , more than 2 times..

                                    if (tds.Tables[0].Columns[userRoleColumnCount].ColumnName == "UserRole" && Convert.ToString(drtableOld.ItemArray[userRoleColumnCount]) != "")
                                    {
                                        userRoleDataTable.ImportRow(drtableOld);
                                    }
                                    else if (value == "AuditScenario")
                                    {
                                        auditScenarioTable.ImportRow(drtableOld);
                                    }
                                    else if (value == "Scenario")
                                    {
                                        scenarioDataTable.ImportRow(drtableOld);
                                    }
                                    else if (value == "CurrencyCondition")
                                    {
                                        currencyConditionDataTable.ImportRow(drtableOld);
                                    }
                                    else if (value == "InputType")
                                    {
                                        inputTypeDataTable.ImportRow(drtableOld);
                                    }
                                    else if (value == "Interval")
                                    {
                                        intervalDataTable.ImportRow(drtableOld);
                                    }
                                    else if (value == "ProductLine")
                                    {
                                        productLineDataTable.ImportRow(drtableOld);
                                    }
                                    //added by mounika
                                    else if (value == "Country")
                                    {
                                        countryDataTable.ImportRow(drtableOld);
                                    }
                                    //added by mounika
                                    else if (value == "DeviceType")
                                    {
                                        deviceTypeDataTable.ImportRow(drtableOld);
                                    }
                                }
                                userRoleDataTable.TableName = "UserRole";
                                scenarioDataTable.TableName = "Scenario";
                                currencyConditionDataTable.TableName = "Currency";
                                inputTypeDataTable.TableName = "InputType";
                                intervalDataTable.TableName = "Interval";
                                auditScenarioTable.TableName = "AuditScenario";
                                productLineDataTable.TableName = "ProductLine";
                                //added by mounika
                                countryDataTable.TableName = "Country";
                                deviceTypeDataTable.TableName = "Device Type";

                                localDataSet.Tables.Add(scenarioDataTable);
                                localDataSet.Tables.Add(currencyConditionDataTable);
                                localDataSet.Tables.Add(inputTypeDataTable);
                                localDataSet.Tables.Add(intervalDataTable);
                                localDataSet.Tables.Add(userRoleDataTable);
                                localDataSet.Tables.Add(auditScenarioTable);
                                localDataSet.Tables.Add(productLineDataTable);
                                //added by mounika
                                localDataSet.Tables.Add(countryDataTable);
                                localDataSet.Tables.Add(deviceTypeDataTable);

                                if (tds.Tables.Contains(clsInformation.promoOperations))
                                {
                                    int indexOfOperationsTable = tds.Tables.IndexOf(clsInformation.promoOperations);
                                    promoOperationsTable = tds.Tables[indexOfOperationsTable].Clone();
                                    promoOperationsTable.TableName = clsInformation.promoOperations;

                                    foreach (DataRow row in tds.Tables[indexOfOperationsTable].Rows)
                                    {
                                        promoOperationsTable.ImportRow(row);
                                    }

                                    //DataTable reversedRows = ReverseRowsInDataTable(promoOperationsTable);

                                    DataRow selectedRow = promoOperationsTable.Rows[2];
                                    DataRow newRow = promoOperationsTable.NewRow();
                                    newRow.ItemArray = selectedRow.ItemArray;
                                    promoOperationsTable.Rows.Remove(selectedRow);
                                    promoOperationsTable.Rows.InsertAt(newRow, 4);

                                    localDataSet.Tables.Add(ReverseRowsInDataTable(promoOperationsTable));
                                }
                            }
                        }
                    }

                    break;

                case "Upload":
                default:
                    if (FAST.response != "")
                    {
                        using (StringReader stringReader = new StringReader(FAST.response))
                        {
                            localDataSet.ReadXml(stringReader);
                        }
                    }
                    break;

                case "Refresh Branson Data":
                    FAST.response = "<RefreshBranson>" + FAST.response + "</RefreshBranson>";
                    using (StringReader stringReader = new StringReader(FAST.response))
                    {
                        localDataSet.ReadXml(stringReader);
                    }
                    break;

            }

            return localDataSet;
        }


        private static DataTable ReverseRowsInDataTable(DataTable inputTable)
        {
            DataTable outputTable = inputTable.Clone();

            for (int i = inputTable.Rows.Count - 1; i >= 0; i--)
            {
                outputTable.ImportRow(inputTable.Rows[i]);
            }

            return outputTable;
        }
        #endregion


    }





}














