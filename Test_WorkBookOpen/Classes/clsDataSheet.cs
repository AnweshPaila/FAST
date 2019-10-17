using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace Test_WorkBookOpen.Classes
{
    class clsDataSheet
    {
        #region  buildSheet
        /// <summary>
        /// Calling this method to generate the Reference Sheet Data
        /// </summary>
        /// <param name="worksheet">Reference sheet information will sent as parameter to this method.</param>
        public static void buildSheet(Excel.Worksheet worksheet)
        {

            if (worksheet == null && worksheet.Name != clsInformation.referenceDataSheet)
                throw new Exception("The Sheet is Not Validated For Data Building");

            worksheet.UsedRange.Clear();
            populateAllFilters(worksheet);

        }
        #endregion

        #region populateAllFilters
        /// <summary>
        /// This method is used to add data to the sheet for the Dropdown Filters
        /// </summary>
        /// <param name="worksheet">Reference sheet information will sent as parameter to this method.</param>
        public static void populateAllFilters(Excel.Worksheet worksheet)
        {

            // Starting RowIndex For Refdata Body
            int startRowIndex = 2;
            generateRefDataFirstRowHeaders(1, worksheet);

            #region For Building RefsheetData
            // For process Dropdown data
            generateRefDataBody(1, 1, startRowIndex, 0, worksheet, 1);
            // For InputType Dropdown data
            generateRefDataBody(1, 3, startRowIndex, 2, worksheet, 2);

            #endregion

        }
        #endregion

        #region Sheet Headers and Body

        /// <summary>
        /// This Method is used to generate the First Row Information(Headers) 
        /// </summary>
        /// <param name="columnValue">which column to be considered for binding the data</param>
        /// <param name="workSheet">Reference sheet information will sent as parameter to this method.</param>
        public static void generateRefDataFirstRowHeaders(int columnValue, Excel.Worksheet workSheet)
        {
            int startRowValue = 1;
            workSheet.UsedRange.Clear();

            for (int i = 0; i < 2; i++)
            {

                // Applicable for All Types
                workSheet.Cells[startRowValue, columnValue].Value = clsInformation.headersForReferenceSheet[0];
                workSheet.Cells[startRowValue, ++columnValue].Value = clsInformation.headersForReferenceSheet[1];

                if (i == 1)
                {
                    // Applicable for InputType
                    workSheet.Cells[startRowValue, columnValue + 1].Value = clsInformation.headersForReferenceSheet[2];
                    workSheet.Cells[startRowValue, columnValue + 2].Value = clsInformation.headersForReferenceSheet[3];
                    workSheet.Cells[startRowValue, columnValue + 3].Value = clsInformation.headersForReferenceSheet[4];
                    workSheet.Cells[startRowValue, columnValue + 4].Value = clsInformation.headersForReferenceSheet[5];
                    workSheet.Cells[startRowValue, columnValue + 5].Value = clsInformation.headersForReferenceSheet[6];
                    workSheet.Cells[startRowValue, columnValue + 6].Value = clsInformation.headersForReferenceSheet[7];
                    workSheet.Cells[startRowValue, columnValue + 7].Value = clsInformation.headersForReferenceSheet[8];
                    workSheet.Cells[startRowValue, columnValue + 8].Value = clsInformation.headersForReferenceSheet[9];

                    columnValue = columnValue + 8;
                }

                columnValue++;
            }

            for (int i = 0; i < clsInformation.headernames.Length; i++)
            {
                workSheet.Cells[1, columnValue].Value = clsInformation.headernames[i];
                columnValue++;
            }

        }

        /// <summary>
        /// This Method is used to generate the Reference sheet Body
        /// </summary>
        /// <param name="rowValue">Indidcates from which row, the data has to be binded</param>
        /// <param name="colValue">Indidcates from which column, the data has to be binded</param>
        /// <param name="startRowIndex">Starting Row Number is sent as parameter</param>
        /// <param name="tableid">The table ID where the application needs to find data</param>
        /// <param name="worksheet">Reference Sheet Information is sent as parameter</param>
        /// <param name="mode">Indicates which mode needs to be executed</param>
        public static void generateRefDataBody(int rowValue, int colValue, int startRowIndex, int tableid, Excel.Worksheet worksheet, int mode)
        {
            // For Adding Initialize Workbook data to the Refdata sheet
            if (mode == 1)
            {
                for (int i = 0; i < FAST._dsInitilaizeWorkbook.Tables[tableid].Rows.Count; i++)
                {
                    worksheet.Cells[startRowIndex, colValue].Value = Convert.ToInt32(FAST._dsInitilaizeWorkbook.Tables[tableid].Rows[i].ItemArray[1]);
                    worksheet.Cells[startRowIndex, colValue + 1].Value = Convert.ToString((FAST._dsInitilaizeWorkbook.Tables[tableid].Rows[i].ItemArray[0]));

                    startRowIndex++;
                }
            }
            else
            {

                for (int i = 0; i < FAST._dsAllFilters.Tables[tableid].Rows.Count; i++)
                {

                    if (Convert.ToString(FAST._dsAllFilters.Tables[tableid].Rows[i].ItemArray[1]) == clsInformation.scenario ||
                        Convert.ToString(FAST._dsAllFilters.Tables[tableid].Rows[i].ItemArray[1]) == clsInformation.currencyCondition)
                    {
                        worksheet.Cells[startRowIndex, colValue].Value = Convert.ToInt32(FAST._dsAllFilters.Tables[tableid].Rows[i].ItemArray[5]); // For Id
                        worksheet.Cells[startRowIndex, colValue + 1].Value = Convert.ToString((FAST._dsAllFilters.Tables[tableid].Rows[i].ItemArray[2])); // For Value
                        worksheet.Cells[startRowIndex, colValue + 2].Value = Convert.ToString((FAST._dsAllFilters.Tables[tableid].Rows[i].ItemArray[1])); // For Type
                    }
                    else if (Convert.ToString(FAST._dsAllFilters.Tables[tableid].Rows[i].ItemArray[1]) == clsInformation.inputType)
                    {
                        worksheet.Cells[startRowIndex, colValue].Value = Convert.ToInt32(FAST._dsAllFilters.Tables[tableid].Rows[i].ItemArray[5]); // For Id
                        worksheet.Cells[startRowIndex, colValue + 1].Value = Convert.ToString((FAST._dsAllFilters.Tables[tableid].Rows[i].ItemArray[2])); // For Value
                        worksheet.Cells[startRowIndex, colValue + 2].Value = Convert.ToString((FAST._dsAllFilters.Tables[tableid].Rows[i].ItemArray[1])); // For Type
                        worksheet.Cells[startRowIndex, colValue + 3].Value = Convert.ToInt32(FAST._dsAllFilters.Tables[tableid].Rows[i].ItemArray[0]); // For AllFieldsRequired
                        worksheet.Cells[startRowIndex, colValue + 4].Value = Convert.ToString(Convert.ToDouble(FAST._dsAllFilters.Tables[tableid].Rows[i].ItemArray[6])); // For Minimum Value
                        worksheet.Cells[startRowIndex, colValue + 5].Value = Convert.ToString(Convert.ToDouble(FAST._dsAllFilters.Tables[tableid].Rows[i].ItemArray[3])); // For Maximum Value
                        worksheet.Cells[startRowIndex, colValue + 6].Value = Convert.ToString((FAST._dsAllFilters.Tables[tableid].Rows[i].ItemArray[7])); // For InputTypeDescription
                        worksheet.Cells[startRowIndex, colValue + 7].Value = Convert.ToString((FAST._dsAllFilters.Tables[tableid].Rows[i].ItemArray[8])); // For InputDataType
                        worksheet.Cells[startRowIndex, colValue + 8].Value = Convert.ToString((FAST._dsAllFilters.Tables[tableid].Rows[i].ItemArray[4])); // For VarianceFlagType
                        worksheet.Cells[startRowIndex, colValue + 9].Value = Convert.ToString((FAST._dsAllFilters.Tables[tableid].Rows[i].ItemArray[9])); // For VariancePercentage
                       

                    }

                    startRowIndex++;
                }
            }

        }

        #endregion

        #region For Offline
        /// <summary>
        /// This method is used to bind the data to reference, that is required for the Offline
        /// </summary>
        /// <param name="workSheet">Sheet to which the data has to be binded is sent as parameter</param>
        public static void addDataforOffline(Excel.Worksheet workSheet)
        {
            if (workSheet == null && workSheet.Name != clsInformation.referenceDataSheet)
                throw new Exception("The Sheet is Not Validated For Data Building");


            
            if (FAST._dsDownloadData.Tables.Count > 2)
            {
                FAST._userRole = Convert.ToString(FAST._dsDownloadData.Tables[0].Rows[0].ItemArray[1]); // For UserRole
                FAST._readOnly = Convert.ToByte(FAST._dsDownloadData.Tables[0].Rows[0].ItemArray[0]); // For Readonly
            }
            if (FAST._dsDownloadData.Tables[3].Rows.Count != 0)
            {
                FAST._scenarioStatus = Convert.ToString((FAST._dsDownloadData.Tables[3].Rows[0].ItemArray[4])); // For Sccenario Status

                if (!string.IsNullOrEmpty(Convert.ToString((FAST._dsDownloadData.Tables[3].Rows[0].ItemArray[4]))))
                {
                    FAST._readOnlyStartMonth = Convert.ToString((FAST._dsDownloadData.Tables[3].Rows[0].ItemArray[5])); // For RO-SM

                }

                if (!string.IsNullOrEmpty(Convert.ToString((FAST._dsDownloadData.Tables[3].Rows[0].ItemArray[0]))))
                {
                    FAST._readOnlyEndMonth = Convert.ToString((FAST._dsDownloadData.Tables[3].Rows[0].ItemArray[1])); // For RO-EM

                }

                // storing that value to the variable here
                FAST._startDateTimeStamp = Convert.ToDateTime(FAST._dsDownloadData.Tables[3].Rows[0].ItemArray[3]);// For StartMonthTimeStamp
                FAST._endDateTimeStamp = Convert.ToDateTime(FAST._dsDownloadData.Tables[3].Rows[0].ItemArray[2]); // For EndMonthTimeStamp
            }


        }

        /// <summary>
        /// Used to bind some more data, that is required for Upload Button click in Offline
        /// </summary>
        /// <param name="worksheet"></param>
        public static void uploadButtonClickforOffline(Excel.Worksheet worksheet)
        {
            int startRowIndex = 2, colValue = 13;
            if (worksheet == null && worksheet.Name != clsInformation.referenceDataSheet)
                throw new Exception("The Sheet is Not Validated For Data Building");

            worksheet.Cells[startRowIndex, colValue++] = FAST._valueProcess; // For ProcessId for Upload during Offline
            worksheet.Cells[startRowIndex, colValue++] = FAST._valueScenario; // For ScenarioId
            worksheet.Cells[startRowIndex, colValue++] = FAST._valueInputType; // For InputTypeId
            worksheet.Cells[startRowIndex, colValue++] = FAST._valueCurrency; // For CurrencyConditionId
            worksheet.Cells[startRowIndex, colValue++] = FAST._valueInterval; // For valueInterval
            worksheet.Cells[startRowIndex, colValue++] = FAST._valueProductLine; // For valueInterval //Added by Sita
            worksheet.Cells[startRowIndex, colValue++] = FAST._downloadedProcessValueForOfflineOnline;
            worksheet.Cells[startRowIndex, colValue++] = FAST._downloadedScenarioValueForOfflineOnline;
            worksheet.Cells[startRowIndex, colValue++] = FAST._downloadedInputTypeValueForOfflineOnline;
            worksheet.Cells[startRowIndex, colValue++] = FAST._downloadedCurrencyValueForOfflineOnline;
            worksheet.Cells[startRowIndex, colValue++] = FAST._downloadedIntervalValueForOfflineOnline;
            worksheet.Cells[startRowIndex, colValue++] = FAST._downloadedProductLineValueForOfflineOnline;
            worksheet.Cells[startRowIndex, colValue++] = FAST._minValue; // For MinValue
            worksheet.Cells[startRowIndex, colValue++] = FAST._maxValue; // For MaxValue
            worksheet.Cells[startRowIndex, colValue++] = FAST._allFieldsRequiredConditionForUpload; // All Fields Required or not condition for Upload
            worksheet.Cells[startRowIndex, colValue++] = FAST._dataTypeValue; // Used for checking the datatype
            worksheet.Cells[startRowIndex, colValue++] = FAST._saveDataFile; // Used for DataTableSaveFile 
            worksheet.Cells[startRowIndex, colValue++] = FAST._startRange; // Used for Start Range 
            worksheet.Cells[startRowIndex, colValue++] = FAST._endRange; // Used for End Range
            worksheet.Cells[startRowIndex, colValue++] = clsManageSheet.bodyRowStartingNumber;
            worksheet.Cells[startRowIndex, colValue++] = clsManageSheet.formulaNextColumn;
            worksheet.Cells[startRowIndex, colValue++] = FAST._txtProcess;

            //anwesh 24//08/2019
            worksheet.Cells[(startRowIndex - 1), colValue++] = clsInformation.AliasId;
            worksheet.Cells[startRowIndex, (colValue-1)] = FAST.userName;

        }
        
        #endregion

        public static void WriteToLogFile(string Message)
        {
            //string FilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), BusinessConstantFields.CONST_ROYALTY_CALCULATION_SYSTEM.Replace("_", "")).ToString();
            string FilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)).ToString();
            
            FilePath = Path.Combine(FilePath, "log.txt").ToString();

            StreamWriter file = new StreamWriter(FilePath, true);
            file.WriteLine(Message);
            file.Close();
        }


    }
}
