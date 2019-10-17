using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Microsoft.Office.Tools.Excel;
using System.Data;
using System.Drawing;

namespace Test_WorkBookOpen.Classes
{
    class clsManageSheet
    {
        public static string[] fields;
        public static int inputTemplateRowBodyNumber, varianceReportrowBodyNumber, bodyRowStartingNumber, varianceReporMonthlyCol, inputTemplateMonthlyStartColumnNumber, inputTemplateMonthlyLastColumnNumber;
        public static string formulaColumnName, formulaNextColumn, varianceReportPCLastColumn, varianceReportMonthlyData, currencyColumnName, countryColumName;

        #region BuildSheet
        /// <summary>
        /// This Method is used to generate sheets in the opened Workbook
        /// </summary>
        /// <param name="sheet">Generated sheet Info is passed back</param>
        /// <param name="sheetName">Sheet Name is sent as parameter for creating or generating the sheet in the current Workbook</param>
        public static void buildSheet(ref Excel.Worksheet sheet, string sheetName)
        {
            if (string.IsNullOrEmpty(sheetName))
                throw new Exception("Parameters are Invalid for Sheet Creation");

            Workbook wrkbk = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook);

            // storing the count of sheets available in the active workbook

            int wrkbkAvaialableSheetCount = wrkbk.Worksheets.Count;

            // iterating the count of the sheetscount

            for (int i = 1; i <= wrkbkAvaialableSheetCount; i++)
            {
                // Getting Each Sheet Data into the variable declared below

                Excel.Worksheet wrksht = (Excel.Worksheet)wrkbk.Worksheets.get_Item(i);

                // Comparing the WorkSheetName with the HiddenSheetName Passed to the Method

                if (wrksht.Name == sheetName)
                {
                    sheet = wrksht;

                    return;
                }
            }
            sheet = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(Type.Missing, Type.Missing, 1, Type.Missing) as Excel.Worksheet;

            sheet.Name = sheetName;
        }
        #endregion

        #region BuildSheetTitle

        /// <summary>
        /// This method is used to generate the Title Information for the Sheet
        /// </summary>
        /// <param name="sheet">Information of the sheet on which Title has to be created</param>
        /// <param name="process">What process the user has selected</param>
        /// <param name="scenario">What Scenario the user has selected</param>
        /// <param name="inputtype">What InputType the user has selected</param>
        /// <param name="currency">What Currency the user has selected</param>
        /// <param name="previousscenario">This Parameter is used for generation of variance report</param>
        /// <param name="productLine">Product Line selected by the user</param>

        public static void buildSheetTitle(string sheetName, string process, string scenario, string inputtype, string currency, string previousscenario, string intervel, string productLine)
        {
            FAST.updateEvents(false);

            if (sheetName != null && sheetName != "")
            {

                // getting the data of Product Revenue Work Sheet in this case
                Worksheet sheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[sheetName]);

                if (sheet == null)
                {
                    throw new Exception("Worksheet cannot be empty when building sheet body");
                }

                if (sheet.Visible == Excel.XlSheetVisibility.xlSheetVisible)
                    sheet.Activate();

                unProtect(sheet); //calling the UnProtect Method to remove the protection to the sheet

                clsInformation.sheetTitleContinueRowNumber = 5;

                if (sheet.Name != null)
                {
                    Excel.Range sheetTitleRange = sheet.Range[clsInformation.constSheetTitleStartRange, clsInformation.constSheetTitleEndRange] as Excel.Range;
                    sheetTitleRange.Merge(false);                                        // Here we are not Merging the cells for Sheet Title
                    sheetTitleRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; // Aligining the text for the cells to be in center
                    sheetTitleRange.VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;  // Aligining the text for the cells to be in center
                    sheetTitleRange.ColumnWidth = 24;
                    sheetTitleRange.Value = clsInformation.sheetTitle;
                    sheetTitleRange.Font.Size = 14;
                    sheetTitleRange.Interior.Color = ColorTranslator.FromHtml(clsInformation.colorBlack); // For Color Translator Using System.Drawing Namespace #006599
                    sheetTitleRange.Font.ColorIndex = 2; // For White Text, 1 for Black
                    sheetTitleRange.Font.Bold = true;

                    Excel.Range shtTitleRange1 = sheet.Range[clsInformation.constSheetSummaryStartRange, clsInformation.constSheetSummaryEndRange] as Excel.Range;
                    shtTitleRange1.Merge(false);
                    shtTitleRange1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    shtTitleRange1.VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    shtTitleRange1.Value = clsInformation.filterSummary;
                    shtTitleRange1.Interior.Color = ColorTranslator.FromHtml(clsInformation.colorOrange);
                    shtTitleRange1.Font.ColorIndex = 1; // For White Text, 1 for Black ----
                    shtTitleRange1.Font.Bold = true;


                    sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Value = clsInformation.process;
                    sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Font.Bold = true;
                    sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber + 1].Value = process;

                    sheet.Cells[++clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Value = clsInformation.scenario;
                    sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Font.Bold = true;
                    sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber + 1].Value = scenario;

                    #region Graphs and Scenario
                    if (sheet.Name == clsInformation.statistics)
                    {

                        sheet.Cells[++clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Value = clsInformation.user;
                        sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Font.Bold = true;
                        sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber + 1].Value = FAST.userName;
                        sheet.Cells[++clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Value = clsInformation.downloadTimeStamp;
                        sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Font.Bold = true;
                        sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber + 1].Value = DateTime.Now;
                        sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber + 1].NumberFormat = ("MMM dd, yyyy h:mm AM/PM");

                        inputTemplateRowBodyNumber = clsInformation.sheetTitleContinueRowNumber + 2;
                        splitWindow(clsInformation.statistics, 0, 5);
                    }
                    #endregion

                    #region Audit Reports and Not Graphs

                    if (sheet.Name != clsInformation.productsAuditReport && sheet.Name != clsInformation.statistics && sheet.Name != clsInformation.productsAuditReportPivot)
                    {
                        sheet.Cells[++clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Value = clsInformation.inputType;
                        sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Font.Bold = true;
                        sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber + 1].Value = inputtype;

                        string minValue = "", maxValue = "";

                        if (FAST._dataTypeValue == clsInformation.decimalType)
                        {
                            if (Convert.ToString(FAST._minValue).Contains("."))
                            {
                                minValue = Convert.ToString(FAST._minValue);
                            }
                            else if (!Convert.ToString(FAST._minValue).Contains("."))
                            {
                                minValue = (FAST._minValue).ToString("#,##0.00");
                            }


                            if (Convert.ToString(FAST._maxValue).Contains("."))
                            {
                                maxValue = Convert.ToString((FAST._maxValue));
                            }
                            else if (!Convert.ToString(FAST._maxValue).Contains("."))
                            {
                                maxValue = (FAST._maxValue).ToString("#,##0.00");
                            }

                        }
                        else
                        {
                            // Added for InputType Min Value and Max Value
                            minValue = Convert.ToString(Convert.ToDecimal(Math.Round(FAST._minValue * 100)));
                            maxValue = Convert.ToString(Convert.ToDecimal(Math.Round(FAST._maxValue * 100)));
                        }


                        sheet.Cells[++clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Value = "Description";
                        sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Font.Bold = true;
                        sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber + 1].Value = FAST._description;
                        
                        if (FAST._valueInputType != "0")
                        {
                            sheet.Cells[++clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Value = "Min Value";
                            sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber + 1].Value = minValue;
                            sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Font.Bold = true;
                            sheet.Cells[++clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Value = "Max Value";
                            sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber + 1].Value = maxValue;
                            sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Font.Bold = true;
                        }

                        if (FAST._valueInputType == "0")
                        {
                            sheet.Cells[++clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Value = "Min Value";
                            sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber + 1].Value = "N/A";
                            sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Font.Bold = true;
                            sheet.Cells[++clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Value = "Max Value";
                            sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber + 1].Value = "N/A";
                            sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Font.Bold = true;
                        }

                        sheet.Cells[++clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Value = "DataType";
                        sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Font.Bold = true;
                        sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber + 1].Value = FAST._dataTypeValue;

                    }
                    else if (sheet.Name != clsInformation.statistics)
                    {
                        sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Value = clsInformation.auditScenario;


                        sheet.Cells[++clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Value = clsInformation.user;
                        sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Font.Bold = true;
                        sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber + 1].Value = FAST.userName;
                        sheet.Cells[++clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Value = clsInformation.downloadTimeStamp;
                        sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Font.Bold = true;
                        sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber + 1].Value = DateTime.Now;
                        sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber + 1].NumberFormat = ("MMM dd, yyyy h:mm AM/PM");


                        inputTemplateRowBodyNumber = clsInformation.sheetTitleContinueRowNumber + 2;
                        splitWindow(sheet.Name, inputTemplateRowBodyNumber, 0);

                    }
                    #endregion

                    #region Variance Report
                    if (previousscenario != "" && previousscenario != null && sheet.Name != clsInformation.statistics)
                    {

                        sheet.Cells[++clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Value = clsInformation.previousScenario;
                        sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Font.Bold = true;
                        sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber + 1].Value = previousscenario;

                        //Taken off Currency from the sheet title
                        //sheet.Cells[++clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Value = clsInformation.currency;
                        //sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Font.Bold = true;
                        //sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber + 1].Value = currency;

                        sheet.Cells[++clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Value = clsInformation.user;
                        sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Font.Bold = true;
                        sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber + 1].Value = FAST.userName;
                        sheet.Cells[++clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Value = clsInformation.downloadTimeStamp;
                        sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Font.Bold = true;
                        sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber + 1].Value = DateTime.Now;
                        sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber + 1].NumberFormat = ("MMM dd, yyyy h:mm AM/PM");

                        varianceReportrowBodyNumber = clsInformation.sheetTitleContinueRowNumber + 3;
                        splitWindow(clsInformation.productsRevenueReport, varianceReportrowBodyNumber, 0);



                    }
                    else if (sheet.Name != clsInformation.productsAuditReport && sheet.Name != clsInformation.statistics && sheet.Name != clsInformation.productsAuditReportPivot)
                    {
                        if (process == "Accounting View")
                        {

                            sheet.Cells[++clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Value = clsInformation.currency;
                            sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Font.Bold = true;
                            sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber + 1].Value = currency;
                            sheet.Cells[++clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Value = clsInformation.user;
                            sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Font.Bold = true;
                            sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber + 1].Value = FAST.userName;
                            sheet.Cells[++clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Value = clsInformation.downloadTimeStamp;
                            sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Font.Bold = true;
                            sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber + 1].Value = DateTime.Now;
                            sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber + 1].NumberFormat = ("MMM dd, yyyy h:mm AM/PM");
                        }
                        else
                        {

                            sheet.Cells[++clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Value = "Interval";
                            sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Font.Bold = true;
                            sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber + 1].Value = intervel;

                            sheet.Cells[++clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Value = "Product Line";
                            sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Font.Bold = true;
                            sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber + 1].Value = productLine;

                            //sheet.Cells[++clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Value = clsInformation.currency;
                            //sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Font.Bold = true;
                            //sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber + 1].Value = currency;
                            sheet.Cells[++clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Value = clsInformation.user;
                            sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Font.Bold = true;
                            sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber + 1].Value = FAST.userName;
                            sheet.Cells[++clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Value = clsInformation.downloadTimeStamp;
                            sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber].Font.Bold = true;
                            sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber + 1].Value = DateTime.Now;
                            sheet.Cells[clsInformation.sheetTitleContinueRowNumber, clsInformation.sheetTitleContinueColumnNumber + 1].NumberFormat = ("MMM dd, yyyy h:mm AM/PM");
                        }

                        inputTemplateRowBodyNumber = clsInformation.sheetTitleContinueRowNumber + 2;
                        splitWindow(clsInformation.productRevenue, inputTemplateRowBodyNumber, 0);

                    }
                    #endregion

                    sheet.Range["C5", "C" + clsInformation.sheetTitleContinueRowNumber].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                    if (sheet.Name == clsInformation.productRevenue)
                        bodyRowStartingNumber = inputTemplateRowBodyNumber;

                }
                else
                    throw new Exception("Sheet object can not be empty when creating the sheet title");


                protectSheet(sheet);

            }

            FAST.updateEvents(true);

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheetName"></param>
        private static void splitWindow(string sheetName, int splitRow, int splitcolumn)
        {
            if (sheetName != clsInformation.statistics && sheetName != clsInformation.productsAuditReport)
            {
                Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[sheetName].Application.ActiveWindow.SplitRow = splitRow;
                Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[sheetName].Application.ActiveWindow.FreezePanes = true;
                Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[sheetName].Application.ActiveWindow.Zoom = 80;
            }
            else if (sheetName == clsInformation.statistics || sheetName == clsInformation.productsAuditReportPivot || sheetName == clsInformation.productsAuditReport)
            {
                Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[sheetName].Application.ActiveWindow.Zoom = 80;
            }

        }
        #endregion

        #region buildSheetBody

        /// <summary>
        /// This Method is used to build the Sheet Body
        /// </summary>
        /// <param name="shtName">Name of the sheet on which the body has to be generated</param>
        /// <param name="lo">Listobject useful to bind the data</param>
        /// <param name="dtSource">The source from where the data to be displayed</param>

        public static void buildSheetBody(string shtName, ref ListObject lo, DataTable dtSource, string process, string scenario, string inputtype, string currency, string previousscenario, string intervel, string productLine)
        {
            int rowCnt = dtSource.Rows.Count;

            string columnNames = null;

            columnNames = Convert.ToString(FAST._dsDownloadData.Tables[1].Rows[0].ItemArray[0]).Replace("isReadOnly,", string.Empty);

            fields = columnNames.Split(',');

            for (int i = 0; i < fields.Length; i++)
            {
                fields[i] = fields[i].Replace("_x0020_", " ");

                if (fields[i] == " ")
                {
                    formulaColumnName = getColumnName(i + 2);
                    formulaNextColumn = getColumnName(i + 3);
                    inputTemplateMonthlyStartColumnNumber = i + 3;
                }
                // For Currency Column Name
                if (fields[i] == clsInformation.currency)
                {
                    currencyColumnName = getColumnName(i + 2);
                }
                // For CountryColumnName
                if (fields[i] == clsInformation.country)
                {
                    countryColumName = getColumnName(i + 2);
                }
            }


            inputTemplateMonthlyLastColumnNumber = fields.Length + 1;

            FAST._lastColumnName = getColumnName(inputTemplateMonthlyLastColumnNumber);

            Worksheet worksheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[shtName]);
            if (worksheet == null)
                throw new Exception("Worksheet can not be empty when building sheet body");

            unProtect(worksheet);

            if (lo != null)
                worksheet.Controls.Remove(lo);

            for (int i = 0; i < Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[shtName].ListObjects.Count; i++)
                Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[shtName].ListObjects.Item(i + 1).Delete();


            // Clearing the Used Range
            worksheet.UsedRange.Clear();

            buildSheetTitle(worksheet.Name, process, scenario, inputtype, currency, previousscenario, intervel, productLine);

            unProtect(worksheet);

            FAST.updateEvents(false);

            int rowStartIndex = inputTemplateRowBodyNumber;

            // For Pivot Data
            Excel.Range lstObjRange = worksheet.Range["B" + Convert.ToString(rowStartIndex), (FAST._lastColumnName + Convert.ToString(rowStartIndex + rowCnt))] as Excel.Range;
            lstObjRange.Clear();

            string listObjectName = "ListObject" + shtName;

            lo = worksheet.Controls.AddListObject(lstObjRange, listObjectName);

            for (int i = 0; i < lo.ListColumns.Count; i++)
            {
                lo.ListColumns[i + 1].Name = fields[i];
            }

            BindingSource bs = new BindingSource();
            bs.DataSource = dtSource;
            lo.SetDataBinding(bs, "", fields);

            // This  Method will work only for Accounting View
            if (FAST._txtProcess == "Accounting View")
            {
                modifyHeaders(worksheet);
            }
            else
            {
                modifyTCPUHeaders(worksheet);
            }

            // Adding the Formula Name to the respected Column.
            worksheet.Cells[inputTemplateRowBodyNumber, formulaColumnName].Value = clsInformation.formulaColumnName;


            setSlideDeckListObjectStyle(worksheet, rowStartIndex + 1, dtSource.Rows.Count, FAST._lastColumnName);

            double count = Globals.ThisAddIn.Application.WorksheetFunction.CountIf(lstObjRange, "null");

            while (count >= 1)
            {
                removeNulls(lstObjRange);
                count = Globals.ThisAddIn.Application.WorksheetFunction.CountIf(lstObjRange, "null");
            }

            Excel.Range verifyNullsRange = worksheet.Range[formulaNextColumn + Convert.ToString(rowStartIndex + 1), (FAST._lastColumnName + Convert.ToString(rowStartIndex + rowCnt))] as Excel.Range;
            count = Globals.ThisAddIn.Application.WorksheetFunction.CountIf(verifyNullsRange, "null");

            if (count >= 1)
            {
                removeInputTemplateNulls(dtSource, worksheet);
            }

            lockCorrespondCells(worksheet, dtSource, FAST._lastColumnName);
            FAST.updateEvents(true);

            formatCells(worksheet, dtSource.Rows.Count, FAST._lastColumnName);

            if (FAST._txtProcess == clsInformation.tcpuView)
            {
                lockAndSetLifeTimeValueColumnForMSRP(worksheet.Name);
            }


            FAST.updateEvents(false);

            worksheet.Activate();

            Globals.ThisAddIn.Application.AutoCorrect.AutoFillFormulasInLists = false;

            worksheet.AutoFilterMode = false;

            (worksheet.Range["A1", "A1"] as Excel.Range).ColumnWidth = 5;

            worksheet.Range["B2", "B2"].Select();

            protectSheet(worksheet);

        }


        #endregion

        #region Removing Nulls From the Sheet

        /// <summary>
        /// Removing Nulls from the Given Range of the Sheet
        /// </summary>
        /// <param name="range"></param>

        private static void removeInputTemplateNulls(DataTable dt, Worksheet sheet)
        {

            long lastRowCount = bodyRowStartingNumber + dt.Rows.Count;

            for (long startRowNumber = bodyRowStartingNumber + 1; startRowNumber <= lastRowCount; startRowNumber++)
            {
                for (int startColCount = inputTemplateMonthlyStartColumnNumber; startColCount <= inputTemplateMonthlyLastColumnNumber; startColCount++)
                {
                    if (Convert.ToString(sheet.Cells[startRowNumber, startColCount].Value) == "null")
                    {
                        sheet.Cells[startRowNumber, startColCount].Value = null;
                    }
                }
            }
        }

        private static void removeNulls(Excel.Range range)
        {
            range.Replace(@"null", @""); // Used for Replacing Null Values
        }
        #endregion

        #region Modifying the Sheet Headers

        /// <summary>
        /// Modifying the Header Names 
        /// </summary>
        /// <param name="worksheet"></param>

        private static void modifyHeaders(Worksheet worksheet)
        {
            int colValue = 11, rowStartIndex = inputTemplateRowBodyNumber;


            while (worksheet.Cells[rowStartIndex, colValue].Value != null && !string.IsNullOrEmpty(Convert.ToString(worksheet.Cells[rowStartIndex, colValue].Value)))
            {
                string mnth = Convert.ToString(worksheet.Cells[rowStartIndex, colValue].Value).Replace("'", string.Empty);


                DateTime enteredDate = DateTime.Parse(mnth);

                string[] mnthInfo = mnth.Split(' ');

                string year = "20" + mnthInfo[1];

                var firstDayOfMonth = new DateTime(Convert.ToInt32(year), enteredDate.Month, 1);

                worksheet.Cells[rowStartIndex, colValue].Value = Convert.ToString(firstDayOfMonth.ToString("MM/dd/yyyy"));

                colValue++;

            }


        }

        private static void modifyTCPUHeaders(Worksheet worksheet)
        {
            int colValue = 11, rowStartIndex = inputTemplateRowBodyNumber;


            while (worksheet.Cells[rowStartIndex, colValue].Value != null && !string.IsNullOrEmpty(Convert.ToString(worksheet.Cells[rowStartIndex, colValue].Value)))
            {
                string val = Convert.ToString(worksheet.Cells[rowStartIndex, colValue].Value);

                if (val.Contains("-"))
                {
                    string mnth = val.Replace("-", " ");

                    string[] temp = mnth.Split(' ');

                    mnth = temp[1] + "'" + temp[0];

                    worksheet.Cells[rowStartIndex, colValue].Value = mnth;
                }



                colValue++;

            }




        }
        #endregion

        #region BuildSheetBodyHeader

        /// <summary>
        /// This Method is called when no data is available to bind for the body
        /// </summary>
        /// <param name="shtName">Name of the sheet on which the body has to be generated</param>
        /// <param name="lo">Listobject useful to bind the data</param>

        public static void buildSheetBodyHeader(string shtName, ref ListObject lo, string process, string scenario, string inputtype, string currency, string previousscenario, string intervel, string productLine)
        {

            Worksheet worksheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[shtName]);
            if (worksheet == null)
                throw new Exception("Worksheet can not be empty when building sheet body");

            unProtect(worksheet);

            if (lo != null)
                worksheet.Controls.Remove(lo);

            for (int i = 0; i < Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[shtName].ListObjects.Count; i++)
                Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[shtName].ListObjects.Item(i + 1).Delete();

            worksheet.UsedRange.Clear();

            buildSheetTitle(worksheet.Name, process, scenario, inputtype, currency, previousscenario, intervel, productLine);

            unProtect(worksheet);

            FAST.updateEvents(false);

            int rowStartIndex = inputTemplateRowBodyNumber;

            string[] fields = new string[] { "Product Line","Channel" ,"Country" ,"Program"
                                              ,"Memory Size","Wireless","DTCP","Currency","Formula",
                                              "JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"   };

            FAST._lastColumnName = getColumnName(fields.Length + 1);

            for (int i = 0; i < fields.Length; i++)
            {
                worksheet.Cells[rowStartIndex, i + 2].Value = fields[i];
            }


            setSlideDeckListObjectStyle(worksheet, 0, 0, FAST._lastColumnName);

            worksheet.Activate();

            (worksheet.Range["A1", "A1"] as Excel.Range).ColumnWidth = 5;
            worksheet.Range["B2", "B2"].Select();

            protectSheet(worksheet);


        }

        #endregion

        #region Alternative Coloring


        /// <summary>
        /// Used to apply alternatives colors to the sheetbody
        /// </summary>
        /// <param name="worksheet">Sheet Info is passed here</param>
        /// <param name="startFormulaRowIndex">From which row, the color should be applied</param>
        /// <param name="rowCount">Total Number of rows</param>
        /// <param name="lastColumn">the last column Name</param>

        static void setSlideDeckListObjectStyle(Worksheet worksheet, int startFormulaRowIndex, int rowCount, string lastColumn)
        {
            int startRow = inputTemplateRowBodyNumber;
            Excel.Range formatRng1 = worksheet.Range["B" + Convert.ToString(startRow), lastColumn + Convert.ToString(startRow)] as Excel.Range;
            // #000000 - For Black Color
            formatRng1.Interior.Color = ColorTranslator.FromHtml("#000000");
            // Setting the ColorIndex to 2 for white color.
            formatRng1.Font.ColorIndex = 2;
            formatRng1.Font.Bold = true;
            formatRng1.RowHeight = 20;
            formatRng1.ColumnWidth = 24;
            formatRng1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            formatRng1.VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;


            if (startFormulaRowIndex != 0)
            {
                for (int i = startFormulaRowIndex; i <= startRow + rowCount; i++)
                {

                    // Filters Range
                    Excel.Range rng1 = worksheet.Range["B" + Convert.ToString(i), formulaColumnName + Convert.ToString(i)] as Excel.Range;
                    rng1.Interior.ColorIndex = null;

                    Excel.Range rng2 = worksheet.Range[formulaNextColumn + Convert.ToString(i), (lastColumn + Convert.ToString(i))] as Excel.Range;

                    rng1.Interior.ColorIndex = null;
                    rng2.Interior.ColorIndex = null;

                    #region For Tcpu MSRP (Local Currency)
                    if (FAST._txtProcess == clsInformation.tcpuView)
                    {
                        string msrpValue = Convert.ToString(worksheet.Cells[i, "B"].Value);
                        string countryValue = Convert.ToString(worksheet.Cells[i, countryColumName].Value);


                        if (msrpValue.ToUpper() == clsInformation.msrp.ToUpper() && countryValue.ToUpper() != clsInformation.country.ToUpper() && countryValue.ToUpper() != clsInformation.countryUs.ToUpper())
                        {
                            worksheet.Cells[i, currencyColumnName].Font.Bold = true;
                        }





                        msrpValue = null;
                        countryValue = null;
                    }
                    #endregion

                    if (i % 2 == 0)
                    {
                        rng1.Interior.Color = ColorTranslator.FromHtml("#fdd49b");
                        rng2.Interior.Color = ColorTranslator.FromHtml("#8c8c8c");
                    }
                    else
                    {
                        rng1.Interior.Color = ColorTranslator.FromHtml("#fdfdfd");
                        rng2.Interior.Color = ColorTranslator.FromHtml("#8c8c8c");
                    }
                }


                Excel.Range setColumnWidth = worksheet.Range["D" + Convert.ToString(startFormulaRowIndex), "I" + Convert.ToString(rowCount)] as Excel.Range;
                setColumnWidth.ColumnWidth = 17;

                setColumnWidth = worksheet.Range[formulaNextColumn + Convert.ToString(startFormulaRowIndex), lastColumn + Convert.ToString(rowCount)] as Excel.Range;
                setColumnWidth.ColumnWidth = 25;

                // For Program Column Width
                setColumnWidth = worksheet.Range["E" + startFormulaRowIndex.ToString(), "E" + rowCount.ToString()] as Excel.Range;
                setColumnWidth.ColumnWidth = 27;

            }

        }


        #endregion

        #region Locking corresponding cells

        /// <summary>
        /// This method is used for Protecting some cells from the user from being Modified
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="dtSource"></param>
        /// <param name="lastColumn"></param>

        static void lockCorrespondCells(Worksheet worksheet, DataTable dtSource, string lastColumn)
        {
            int startRowIndex = inputTemplateRowBodyNumber + 1;
            int lastRowCnt = inputTemplateRowBodyNumber + dtSource.Rows.Count;

            if (FAST._readOnly == 0)
            {
                DateTime currentDate = DateTime.Now;
                if (currentDate > Convert.ToDateTime(FAST._startDateTimeStamp) &&
                   currentDate < Convert.ToDateTime(FAST._endDateTimeStamp))
                {

                    Excel.Range monthColumnData = worksheet.Range[formulaNextColumn + startRowIndex, lastColumn + Convert.ToString(lastRowCnt)] as Excel.Range;
                    monthColumnData.Interior.ColorIndex = 19;
                    monthColumnData.Locked = false;

                    // Verifying the Status of the Scenario Dropdown Selected Item Here

                    if (Convert.ToString(FAST._dsDownloadData.Tables[3].Rows[0].ItemArray[4]) == "Open")
                    {

                        //if (FAST._txtProcess == clsInformation.accountingView)
                        //{
                        readOnlyCells(worksheet, dtSource, lastColumn);
                        //}
                        //else if (FAST._txtProcess == clsInformation.tcpu)
                        //{
                        //    lockAndSetLifeTimeValueColumnForMSRP(worksheet, startRowIndex, lastRowCnt, FAST._lastColumnName);
                        //}
                    }
                    else
                    {
                        #region If Scenario Closed Locking All Cells

                        Excel.Range rng = worksheet.Range[formulaNextColumn + startRowIndex, lastColumn + Convert.ToString(lastRowCnt)] as Excel.Range;
                        rng.Interior.Color = ColorTranslator.FromHtml("#8c8c8c");
                        rng.Locked = true;

                        #endregion
                    }

                }
                else
                {

                    if (FAST._userRole == "User")
                    {

                        Excel.Range rng = worksheet.Range[formulaNextColumn + startRowIndex, lastColumn + Convert.ToString(lastRowCnt)] as Excel.Range;
                        rng.Interior.Color = ColorTranslator.FromHtml("#8c8c8c");
                    }
                    else if (FAST._userRole == "SuperUser")
                    {

                        Excel.Range rng = worksheet.Range[formulaNextColumn + startRowIndex, lastColumn + Convert.ToString(lastRowCnt)] as Excel.Range;
                        rng.Interior.ColorIndex = 19;
                        rng.Locked = false;

                        readOnlyCells(worksheet, dtSource, lastColumn);
                    }
                }

            }
            else
            {
                Excel.Range rng = worksheet.Range[formulaNextColumn + startRowIndex, lastColumn + Convert.ToString(lastRowCnt)] as Excel.Range;
                rng.Interior.Color = ColorTranslator.FromHtml("#8c8c8c");
            }

            #region for Unlocking Formula Cell

            // for Formula to be unlocked
            Excel.Range formulaRng = worksheet.Range[formulaColumnName + startRowIndex, formulaColumnName + lastRowCnt] as Excel.Range;
            formulaRng.Locked = false;
            formulaRng.ColumnWidth = 13;

            #endregion

        }

        /// <summary>
        /// This Method is used to get the Column name based on the Number
        /// </summary>
        /// <param name="columnNumber"></param>
        /// <returns></returns>

        public static string getColumnName(int columnNumber)
        {
            const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            string columnName = "";

            while (columnNumber > 0)
            {
                columnName = letters[(columnNumber - 1) % 26] + columnName;
                columnNumber = (columnNumber - 1) / 26;
            }

            return columnName;
        }


        /// <summary>
        /// Ths method is used to grayout the cells and make them readonly
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="dtSource"></param>
        /// <param name="lastcolumn"></param>

        public static void readOnlyCells(Worksheet worksheet, DataTable dtSource, string lastcolumn)
        {

            int startRowIndex = inputTemplateRowBodyNumber;
            int lastRowCnt = inputTemplateRowBodyNumber + dtSource.Rows.Count;
            string startLockrange = null, endLockRange = null;

            //  int colIndex = 11;
            int colIndex = inputTemplateMonthlyStartColumnNumber;

            if (FAST._readOnlyStartMonth != "null" && FAST._readOnlyEndMonth != "null" &&
                FAST._readOnlyStartMonth != null && FAST._readOnlyEndMonth != null &&
                FAST._readOnlyStartMonth != "" && FAST._readOnlyEndMonth != "")
            {
                FAST._readOnlyStartMonth = "01-" + FAST._readOnlyStartMonth.Replace("'", "-").Replace(" ", string.Empty);
                FAST._readOnlyEndMonth = "01-" + FAST._readOnlyEndMonth.Replace("'", "-").Replace(" ", string.Empty);

                DateTime startDate = DateTime.Parse(FAST._readOnlyStartMonth.Replace("'", "-"));
                DateTime endDate = DateTime.Parse(FAST._readOnlyEndMonth.Replace("'", "-"));

                while (worksheet.Cells[startRowIndex, colIndex].Value != null && !string.IsNullOrEmpty(Convert.ToString(worksheet.Cells[startRowIndex, colIndex].Value)))
                {
                    if (Convert.ToString(worksheet.Cells[inputTemplateRowBodyNumber, colIndex].Value) != "Life Time Value" && Convert.ToString(worksheet.Cells[inputTemplateRowBodyNumber, colIndex].Value) != "Column1")
                    {
                        DateTime dt = new DateTime();
                        if (Convert.ToString(worksheet.Cells[inputTemplateRowBodyNumber, colIndex].Value).Contains("'"))
                        {
                            // For TCPU View
                            string date = "01-" + Convert.ToString(worksheet.Cells[inputTemplateRowBodyNumber, colIndex].Value).Replace("'", "-").Replace(" ", string.Empty);
                            dt = Convert.ToDateTime(date);
                        }
                        else
                        {

                            dt = Convert.ToDateTime(Convert.ToString(worksheet.Cells[inputTemplateRowBodyNumber, colIndex].Value));
                        }


                        if (dt == Convert.ToDateTime(startDate.ToString("MM/dd/yyyy")) &&
                            dt == Convert.ToDateTime(endDate.ToString("MM/dd/yyyy")))
                        {
                            // if start month and end month are same
                            startLockrange = getColumnName(colIndex) + Convert.ToString(startRowIndex + 1);
                            endLockRange = getColumnName(colIndex) + Convert.ToString(lastRowCnt);

                            // Useful at the time of upload
                            FAST._startRange = getColumnName(colIndex + 1) + Convert.ToString(startRowIndex + 1);
                            FAST._endRange = lastcolumn + Convert.ToString(lastRowCnt);

                        }
                        else if (dt == Convert.ToDateTime(startDate.ToString("MM/dd/yyyy")))
                        {
                            // For storing the start month Index here
                            startLockrange = getColumnName(colIndex) + Convert.ToString(startRowIndex + 1);
                            FAST._startRange = getColumnName(colIndex) + Convert.ToString(startRowIndex + 1);
                        }
                        else if (dt == Convert.ToDateTime(endDate.ToString("MM/dd/yyyy")))
                        {
                            // For storing the end month Index here
                            endLockRange = getColumnName(colIndex) + Convert.ToString(lastRowCnt);
                            FAST._endRange = getColumnName(colIndex) + Convert.ToString(lastRowCnt);

                            // Useful at the time of upload
                            FAST._startRange = getColumnName(colIndex + 1) + Convert.ToString(startRowIndex + 1);
                            FAST._endRange = lastcolumn + Convert.ToString(lastRowCnt);
                        }

                    }
                    colIndex++;
                }

                if (startLockrange != null && endLockRange != null)
                {
                    Excel.Range rng = worksheet.Range[startLockrange, endLockRange] as Excel.Range;
                    rng.FormatConditions.Delete();
                    rng.Interior.Color = ColorTranslator.FromHtml("#8c8c8c");

                    rng.Locked = true;


                }

                else
                {
                    // When Start range and end range not in between the specified time limit
                    // Useful at the time of upload
                    FAST._startRange = formulaNextColumn + Convert.ToString(startRowIndex + 1);
                    FAST._endRange = lastcolumn + Convert.ToString(lastRowCnt);
                }


            }
            else
            {
                colIndex = 2;

                while (worksheet.Cells[startRowIndex, colIndex].Value != null && !string.IsNullOrEmpty(Convert.ToString(worksheet.Cells[startRowIndex, colIndex].Value)))
                {
                    if (Convert.ToString(worksheet.Cells[startRowIndex, colIndex].Value) == "Formula")
                    {
                        FAST._startRange = getColumnName(colIndex + 1) + Convert.ToString(startRowIndex + 1);
                        FAST._endRange = lastcolumn + Convert.ToString(lastRowCnt);
                    }

                    colIndex++;
                }

            }
        }

        #endregion

        #region Lock Corresponding Column for TCPU
        public static void lockAndSetLifeTimeValueColumnForMSRP(string sheetName)
        {
            Worksheet worksheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[sheetName]);
            if (worksheet == null)
                throw new Exception("Worksheet can not be empty when building sheet body");

            string[] columns = FAST._dsDownloadData.Tables[1].Rows[0].ItemArray[0].ToString().Replace("isReadOnly,", string.Empty).Split(',');

            FAST._lastColumnName = clsManageSheet.getColumnName(columns.Length + 1);

            int isReadOnlyColumnNumber = 0;

            for (int i = 0; i < FAST._dsDownloadData.Tables[2].Columns.Count; i++)
            {
                if (FAST._dsDownloadData.Tables[2].Columns[i].ColumnName == clsInformation.isReadOnly)
                    isReadOnlyColumnNumber = i;
            }


            if (bodyRowStartingNumber != 0)
            {
                for (int i = 0; i < FAST._dsDownloadData.Tables[2].Rows.Count; i++)
                {

                    #region For Tcpu
                    string value = Convert.ToString(FAST._dsDownloadData.Tables[2].Rows[i].ItemArray[isReadOnlyColumnNumber]);

                    if (value == "1") // Making Data Read only
                    {
                        Excel.Range makeReadOnly = worksheet.get_Range(formulaNextColumn + Convert.ToString(bodyRowStartingNumber + i + 1),
                                                                    FAST._lastColumnName + Convert.ToString(bodyRowStartingNumber + i + 1)) as Excel.Range; // adding 1 here as i is starting from 0 which includes the header to remove the header adding 1

                        makeReadOnly.FormatConditions.Delete();

                        makeReadOnly.Interior.Color = null;
                        makeReadOnly.Interior.Color = ColorTranslator.FromHtml("#8c8c8c");
                        makeReadOnly.Locked = true;


                    }

                    value = null;
                }
                #endregion


            }

            // For Hiding the Last Column For TCPU
            //Excel.Range hideRange = worksheet.get_Range(lastColumnForHiding + Convert.ToString("1"), lastColumnForHiding + Convert.ToString("1")) as Excel.Range;
            //hideRange.EntireColumn.Hidden = true;



        }


        #endregion

        #region FormatCells

        /// <summary>
        /// This method is used apply number formation to the binded data
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="rowCount"></param>
        /// <param name="lastColumn"></param>


        public static void formatCells(Worksheet worksheet, int rowCount, string lastColumn)
        {
            if (rowCount <= 0)
                return;
            int startRowIndex = inputTemplateRowBodyNumber;


            // For Pivot Data

            Excel.Range rng = worksheet.Range[formulaNextColumn + Convert.ToString(startRowIndex + 1), lastColumn + Convert.ToString(startRowIndex + rowCount)] as Excel.Range;
            Excel.Range rngFormula = worksheet.Range[(formulaColumnName) + Convert.ToString(startRowIndex + 1), formulaColumnName + Convert.ToString(startRowIndex + rowCount)] as Excel.Range;

            rng.NumberFormat = null;

            if (FAST._dataTypeValue == clsInformation.decimalType)
            {
                // For Formula Column
                rngFormula.NumberFormat = "#,##0.00";
                // titleMinRange.NumberFormat = "#,##0.00";

                // For Months Data
                rng.NumberFormat = "#,##0.00";

                if (FAST._minValue < FAST._maxValue)
                {
                    if (FAST._valueInputType != "0")
                    {

                        rng.Validation.Add(Excel.XlDVType.xlValidateDecimal, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, FAST._minValue, FAST._maxValue);//"999999999999999"

                        //string minValue = Convert.ToString(FAST._minValue).Contains(".00") || !Convert.ToString(FAST._minValue).Contains(".") ? (FAST._minValue).ToString("#,##0") : (FAST._minValue).ToString("#,##0.00");
                        //string maxValue = Convert.ToString(FAST._maxValue).Contains(".00") || !Convert.ToString(FAST._maxValue).Contains(".") ? (FAST._maxValue).ToString("#,##0") : (FAST._maxValue).ToString("#,##0.00");


                        string minValue = "", maxValue = "";

                        if (FAST._dataTypeValue == clsInformation.decimalType)
                        {
                            if (Convert.ToString(FAST._minValue).Contains("."))
                            {
                                minValue = Convert.ToString(FAST._minValue);
                            }
                            else if (!Convert.ToString(FAST._minValue).Contains("."))
                            {
                                minValue = (FAST._minValue).ToString("#,##0.00");
                            }


                            if (Convert.ToString(FAST._maxValue).Contains("."))
                            {
                                maxValue = Convert.ToString((FAST._maxValue));
                            }
                            else if (!Convert.ToString(FAST._maxValue).Contains("."))
                            {
                                maxValue = (FAST._maxValue).ToString("#,##0.00");
                            }
                        }
                        else
                        {
                            minValue = Convert.ToString(Convert.ToDecimal(Math.Round(FAST._minValue * 100)));
                            maxValue = Convert.ToString(Convert.ToDecimal(Math.Round(FAST._maxValue * 100)));
                        }

                        rng.Validation.ErrorMessage = "Please Give inputs between " + minValue + " and " + maxValue;
                    }
                }
                else
                {
                    FAST.displayAlerts(clsInformation.DataValidationCannotSet, 3);
                }
            }
            else if (FAST._dataTypeValue == clsInformation.percentType && FAST._valueInputType != "0")
            {
                // For Formula Column
                rngFormula.NumberFormat = "##0.00%";


                // For Months Data
                rng.NumberFormat = "##0.00%";
                // rng.sty = "Percent";
                rng.Validation.Add(Excel.XlDVType.xlValidateDecimal, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, FAST._minValue, FAST._maxValue);
                rng.Validation.ErrorMessage = "Please Give inputs between " + (Convert.ToDecimal(Math.Round(FAST._minValue * 100)) + "%") + " and " + Convert.ToString((Convert.ToDecimal(Math.Round(FAST._maxValue * 100)) + "%"));


            }

        }
        #endregion

        #region Protect Sheet

        /// <summary>
        /// This Method is used to Protect the Sheet
        /// </summary>
        /// <param name="worksheet"></param>

        public static void protectSheet(Worksheet worksheet)
        {
            worksheet.Protect("InputTemplate", worksheet.ProtectDrawingObjects,
                  true, worksheet.ProtectScenarios, worksheet.ProtectionMode,
                  worksheet.Protection.AllowFormattingCells,
                  worksheet.Protection.AllowFormattingColumns,
                  worksheet.Protection.AllowFormattingRows,
                  worksheet.Protection.AllowInsertingColumns,
                  worksheet.Protection.AllowInsertingRows,
                  worksheet.Protection.AllowInsertingHyperlinks,
                  worksheet.Protection.AllowDeletingColumns,
                  worksheet.Protection.AllowDeletingRows,
                  worksheet.Protection.AllowSorting,
                  true,
                  worksheet.Protection.AllowUsingPivotTables);

        }
        #endregion

        #region Uprotect Sheet

        /// <summary>
        /// This Method is used to unprotect the sheet
        /// </summary>
        /// <param name="visiblesheet"></param>

        public static void unProtect(Worksheet visiblesheet)
        {
            visiblesheet.Unprotect("InputTemplate");
        }
        #endregion

        #region Pivot Sheet

        /// <summary>
        /// This Method is used to build Pivot Report based on the InputTemplate Sheet
        /// </summary>
        /// <param name="shtName">Pivot Sheet Name is passed as parameter</param>
        /// <param name="dtSource"></param>


        public static void buildPivotSheetBody(string shtName, DataTable dtSource)
        {

            Workbook oBook = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook);
            Worksheet pivotSheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[shtName]);
            Worksheet sourceSheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[clsInformation.productRevenue]);

            int dataSouceLength = dtSource.Rows.Count;
            FAST._dataSourceLength = dataSouceLength;
            pivotSheet.UsedRange.Clear();

            if (dataSouceLength != 0)
            {
                int lastRow = inputTemplateRowBodyNumber + dataSouceLength;

                Excel.Range sourceRange = sourceSheet.Range["B" + Convert.ToString(inputTemplateRowBodyNumber), FAST._lastColumnName + lastRow];

                //Creating pivot cache
                var pch = oBook.PivotCaches();

                Excel.PivotCache pc = pch.Create(Excel.XlPivotTableSourceType.xlDatabase, sourceRange);

                // specify first cell for pivot table
                Excel.PivotTable pvt = pc.CreatePivotTable(pivotSheet.Range["B11"], "Input Template Pivot");

                //specify the style to apply to a PivotTable using the PivotTable.TableStyle2 Property
                pvt.SmallGrid = false;
                pvt.TableStyle2 = "PivotStyleMedium9";

                // For Classic Pivot 
                pvt.InGridDropZones = true;
                pvt.RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow);

                // For Number Formate
                if (FAST._dataTypeValue == "Decimal")
                {
                    pivotSheet.Rows.NumberFormat = "#,##0";
                }
                else if (FAST._dataTypeValue == "Percent")
                {
                    pivotSheet.Rows.NumberFormat = "##0.00%";
                }

                //To display the column headers in the PivotTable.
                pvt.ShowTableStyleColumnHeaders = true;

                //To display the row headers in the PivotTable.
                pvt.ShowTableStyleRowHeaders = true;

                //To display the banded columns in the PivotTable.
                pvt.ShowTableStyleColumnStripes = true;

                //To display the banded rows in the PivotTable.
                pvt.ShowTableStyleRowStripes = true;
                //Assigning row filters and column filters

                foreach (Excel.PivotField pf in pvt.PivotFields())
                {

                    if (pf.Name == "JAN,FEB,MAR,APR,MAY,JUN,JUL,AUG,SEP,OCT,NOV,DEC")
                    {

                        pf.Orientation = Excel.XlPivotFieldOrientation.xlRowField;

                    }
                    else if (pf.Name == "Process" || pf.Name == "Scenario" || pf.Name == "InputTemplate" || pf.Name == "Country"
                        || pf.Name == "Channel" || pf.Name == "Wireless" || pf.Name == "Memory" || pf.Name == "Memory Size" || pf.Name == "Formula" ||
                    pf.Name == "DTCP" || pf.Name == "Currency")
                    {
                        pf.Orientation = Excel.XlPivotFieldOrientation.xlPageField;

                    }
                    else if (pf.Name == "Program" || pf.Name == "Product Line" || pf.Name == "Account")
                    {

                        pf.Orientation = Excel.XlPivotFieldOrientation.xlRowField;

                    }
                    else
                    {
                        string caption = pf.Name;
                        pf.Orientation = Excel.XlPivotFieldOrientation.xlDataField;
                        pf.Function = Excel.XlConsolidationFunction.xlSum;
                        pf.Caption = caption + " ";

                        pivotSheet.Rows[11].NumberFormat = ";;;";
                    }
                }


            }

        }

        #endregion

        #region Audit Report Pivot Sheet
        /// <summary>
        /// This Method is used to build Pivot Report based on the InputTemplate Sheet
        /// </summary>
        /// <param name="shtName">Pivot Sheet Name is passed as parameter</param>
        /// <param name="dtSource"></param>


        public static void buildAuditReportPivotSheetBody(string shtName, DataTable dtSource, string process, string scenario)
        {

            Workbook oBook = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook);
            Worksheet pivotSheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[shtName]);
            Worksheet sourceSheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[clsInformation.productsAuditReport]);

            int dataSouceLength = dtSource.Rows.Count;
            FAST._dataSourceLength = dataSouceLength;

            unProtect(pivotSheet);

            pivotSheet.UsedRange.Clear();

            if (dataSouceLength != 0)
            {
                int lastRow = inputTemplateRowBodyNumber + dataSouceLength;

                Excel.Range sourceRange = sourceSheet.Range["B" + Convert.ToString(inputTemplateRowBodyNumber), FAST._lastColumnName + lastRow];

                //Adding title
                buildSheetTitle(shtName, process, scenario, null, null, null, null, null);
                unProtect(pivotSheet);
                //Creating pivot cache
                var pch = oBook.PivotCaches();

                Excel.PivotCache pc = pch.Create(Excel.XlPivotTableSourceType.xlDatabase, sourceRange);

                //Excel.PivotCache pc = pch.Create(Excel.XlPivotTableSourceType.xlExternal, dtSource);

                // specify first cell for pivot table
                Excel.PivotTable pvt = pc.CreatePivotTable(pivotSheet.Range["B" + (inputTemplateRowBodyNumber + 7)], "Audit Report Pivot");

                //specify the style to apply to a PivotTable using the PivotTable.TableStyle2 Property
                pvt.SmallGrid = false;
                pvt.TableStyle2 = "PivotStyleMedium9";

                // For Classic Pivot 
                pvt.InGridDropZones = true;
                pvt.RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow);


                //To display the column headers in the PivotTable.
                pvt.ShowTableStyleColumnHeaders = true;

                //To display the row headers in the PivotTable.
                pvt.ShowTableStyleRowHeaders = true;

                //To display the banded columns in the PivotTable.
                pvt.ShowTableStyleColumnStripes = true;

                //To display the banded rows in the PivotTable.
                pvt.ShowTableStyleRowStripes = true;
                //Assigning row filters and column filters

                foreach (Excel.PivotField pf in pvt.PivotFields())
                {

                    //if(pf.Name == "AliasId")
                    //	pf.Orientation = Excel.XlPivotFieldOrientation.xlRowField;

                    if (pf.Name == "InputType" || pf.Name == "Scenario" || pf.Name == "Process")

                    {
                        pf.Orientation = Excel.XlPivotFieldOrientation.xlPageField;

                    }

                    else if (pf.Name == "AliasId" || pf.Name == "UserName" || pf.Name == "UserAction" || pf.Name == "Account")

                    {

                        pf.Orientation = Excel.XlPivotFieldOrientation.xlRowField;

                    }
                    else if (pf.Name == "AuditTimeStamp")
                    {
                        //string caption = pf.Name;
                        pf.Orientation = Excel.XlPivotFieldOrientation.xlDataField;
                        pf.Function = Excel.XlConsolidationFunction.xlCount;
                        pf.Caption = pf.Name;

                        pivotSheet.Rows[11].NumberFormat = ";;;";
                    }
                }

                 (pivotSheet.Range["B1", "B1"] as Excel.Range).ColumnWidth = 25;
                (pivotSheet.Range["C1", "C1"] as Excel.Range).ColumnWidth = 25;
            }

            pivotSheet.Range["B2", "B2"].Select();
            Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[pivotSheet.Name].Application.ActiveWindow.Zoom = 80;
        }

        #endregion

        #region Refresh Pivot Sheet

        /// <summary>
        /// This Method is used to refresh the Pivot Sheet
        /// </summary>
        /// <param name="shtName"></param>
        /// <param name="dataSourceLength"></param>

        public static void RefreshPivoteGrid(string shtName, int dataSourceLength)
        {
            Workbook oBook = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook);
            oBook.RefreshAll();
        }

        #endregion

        #region Audit ReportBody

        public static void buildAuditSheetBody(string shtName, ref ListObject lo, DataTable dtSource, string process, string scenario)
        {


            int rowCnt = dtSource.Rows.Count;

            string columnNames = Convert.ToString(FAST._dsAuditReport.Tables[2].Rows[0].ItemArray[0]);

            fields = columnNames.Split(',');

            for (int i = 0; i < fields.Length; i++)
            {
                fields[i] = fields[i].TrimStart(' ').TrimEnd(' ');
            }


            int columnCnt = fields.Length + 1;

            FAST._lastColumnName = getColumnName(columnCnt);

            Worksheet worksheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[shtName]);
            if (worksheet == null)
                throw new Exception("Worksheet can not be empty when building sheet body");


            unProtect(worksheet);

            if (lo != null)
                worksheet.Controls.Remove(lo);

            for (int i = 0; i < Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[shtName].ListObjects.Count; i++)
                Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[shtName].ListObjects.Item(i + 1).Delete();


            // Clearing the Used Range
            worksheet.UsedRange.Clear();

            buildSheetTitle(worksheet.Name, process, scenario, null, null, null, null, null);

            unProtect(worksheet);

            FAST.updateEvents(false);

            int rowStartIndex = inputTemplateRowBodyNumber;

            // For Pivot Data

            Excel.Range lstObjRange = worksheet.Range["B" + Convert.ToString(rowStartIndex), (FAST._lastColumnName + Convert.ToString(rowStartIndex + rowCnt))] as Excel.Range;
            lstObjRange.Clear();

            string listObjectName = "ListObject_" + shtName;

            lo = worksheet.Controls.AddListObject(lstObjRange, listObjectName);

            for (int i = 0; i < lo.ListColumns.Count; i++)
            {

                lo.ListColumns[i + 1].Name = fields[i];
            }

            BindingSource bs = new BindingSource();
            bs.DataSource = dtSource;
            lo.SetDataBinding(bs, "", fields);


            applyAlternateColors(worksheet, rowStartIndex + 1, dtSource.Rows.Count, FAST._lastColumnName);

            double count = Globals.ThisAddIn.Application.WorksheetFunction.CountIf(lstObjRange, "null");
            while (count >= 1)
            {
                removeNulls(lstObjRange);
                count = Globals.ThisAddIn.Application.WorksheetFunction.CountIf(lstObjRange, "null");
            }


            FAST.updateEvents(true);

            //Commented as Audit report sheet is hidden always
            //worksheet.Activate();

            Globals.ThisAddIn.Application.AutoCorrect.AutoFillFormulasInLists = false;

            worksheet.AutoFilterMode = false;

            (worksheet.Range["A1", "A1"] as Excel.Range).ColumnWidth = 5;


            protectSheet(worksheet);

        }

        private static void applyAlternateColors(Worksheet worksheet, int startFormulaRowIndex, int rowCount, string lastColumnName)
        {

            int startRow = inputTemplateRowBodyNumber;
            Excel.Range formatRng1 = worksheet.Range["B" + Convert.ToString(startRow), lastColumnName + Convert.ToString(startRow)] as Excel.Range;
            // #000000 - For Black Color
            formatRng1.Interior.Color = ColorTranslator.FromHtml("#000000");
            // Setting the ColorIndex to 2 for white color.
            formatRng1.Font.ColorIndex = 2;
            formatRng1.Font.Bold = true;
            formatRng1.RowHeight = 20;
            formatRng1.ColumnWidth = 24;
            formatRng1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            formatRng1.VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;

            if (startFormulaRowIndex != 0)
            {
                for (int i = startFormulaRowIndex; i <= startRow + rowCount; i++)
                {

                    // Filters Range
                    Excel.Range rng1 = worksheet.Range["B" + Convert.ToString(i), lastColumnName + Convert.ToString(i)] as Excel.Range;
                    rng1.Interior.ColorIndex = null;



                    rng1.Interior.ColorIndex = null;

                    if (i % 2 == 0)
                    {
                        rng1.Interior.Color = ColorTranslator.FromHtml("#fdd49b");

                    }
                    else
                    {
                        rng1.Interior.Color = ColorTranslator.FromHtml("#fdfdfd");

                    }
                }


                Excel.Range setColumnWidth = worksheet.Range["D" + Convert.ToString(startFormulaRowIndex), lastColumnName + Convert.ToString(startFormulaRowIndex)] as Excel.Range;
                setColumnWidth.ColumnWidth = 27;

                Excel.Range setNumberFormat = worksheet.Range[lastColumnName + Convert.ToString(startFormulaRowIndex), lastColumnName + Convert.ToString(startRow + rowCount)] as Excel.Range;
                setNumberFormat.NumberFormat = ("MMM dd, yyyy h:mm AM/PM");


            }


        }


        #endregion

        #region Variance Report

        /// <summary>
        /// This method is used to build variance report body
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="loProductRevenueReport"></param>
        /// <param name="dtSource"></param>
        public static void buildVarianceReportBody(string sheetName, ref ListObject loProductRevenueReport, DataTable dtSource, string process, string scenario, string inputtype, string currency, string previousscenario, string intervel, string productLine)
        {

            #region  Variance Report


            #region Verify Sheet
            Worksheet worksheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[sheetName]);
            if (worksheet == null)
                throw new Exception("Worksheet can not be empty when building sheet body");

            unProtect(worksheet);
            #endregion

            #region Getting Header Names
            string columnNames = Convert.ToString(FAST._dsVarianceReport.Tables[1].Rows[0].ItemArray[1]).Replace("_x0020_", " ");

            fields = columnNames.Split(',');

            int columnCnt = fields.Length + 1;

            string lastColumn = getColumnName(columnCnt);


            for (int i = 0; i < fields.Length; i++)
            {

                if (fields[i].Contains("-") || fields[i].Contains("'") || fields[i].Contains("Life Time Value_"))
                {
                    varianceReportPCLastColumn = getColumnName(i + 1);
                    varianceReportMonthlyData = getColumnName(i + 2);

                    varianceReporMonthlyCol = i + 2;
                    break;
                }

            }

            #endregion

            #region List Object Control Operations
            if (loProductRevenueReport != null)
                worksheet.Controls.Remove(loProductRevenueReport);

            for (int i = 0; i < Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[sheetName].ListObjects.Count; i++)
                Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[sheetName].ListObjects.Item(i + 1).Delete();
            #endregion

            worksheet.UsedRange.Clear();

            #region Sheet Title Part
            buildSheetTitle(worksheet.Name, process, scenario, inputtype, currency, previousscenario, intervel, productLine);
            int rowStartIndex = varianceReportrowBodyNumber;
            int rowCnt = varianceReportrowBodyNumber + dtSource.Rows.Count;
            #endregion


            unProtect(worksheet);

            FAST.updateEvents(false);

            // For Pivot Data
            Excel.Range lstObjRange = worksheet.Range["B" + Convert.ToString(rowStartIndex), lastColumn + Convert.ToString((rowStartIndex + dtSource.Rows.Count + 1))] as Excel.Range;
            lstObjRange.Clear();

            clsInformation.listObjName = "ListObject" + sheetName;
            loProductRevenueReport = worksheet.Controls.AddListObject(lstObjRange, clsInformation.listObjName);

            #region Main Operations
            if (FAST._dataTypeValue == "Decimal")
                lstObjRange.NumberFormat = "#,##0.00";
            else if (FAST._dataTypeValue == "Percent")
                lstObjRange.NumberFormat = "##0.00%";

            var varienceList = new List<string>();
            for (int i = 0; i < loProductRevenueReport.ListColumns.Count; i++)
            {
                loProductRevenueReport.ListColumns[i + 1].Name = fields[i];

                if (fields[i].Contains("Var%"))
                {
                    varienceList.Add(getColumnName(i + 2));

                }
            }


            BindingSource bs = new BindingSource();
            bs.DataSource = dtSource;
            //Bind data to list object.
            loProductRevenueReport.SetDataBinding(bs, "", fields);
            #endregion

            #region Header Modififcation
            //Modifying the Headers for TCPU
            if (FAST._txtProcess != clsInformation.accountingView)
            {
                modifyTCPUVarianceReportMainHeader(worksheet, 11, varianceReportrowBodyNumber);
            }
            #endregion

            #region BackGround Color, Format Conditions
            //Specify font,column width,color and text alignment
            formatVarianceReportColumnsForSelectedRange(worksheet, "B" + Convert.ToString(varianceReportrowBodyNumber - 1), "D" + Convert.ToString(varianceReportrowBodyNumber - 1), true, 17, 2, true, ColorTranslator.FromHtml("#000000")); //Black color
            formatVarianceReportColumnsForSelectedRange(worksheet, "E" + Convert.ToString(varianceReportrowBodyNumber - 1), lastColumn + Convert.ToString(varianceReportrowBodyNumber - 1), true, 17, 2, true, ColorTranslator.FromHtml("#000000"));
            formatVarianceReportColumnsForSelectedRange(worksheet, "D1", varianceReportPCLastColumn + Convert.ToString("1"), 17);
            formatVarianceReportColumnsForSelectedRange(worksheet, varianceReportMonthlyData + Convert.ToString("1"), lastColumn + "1", 17);
            formatVarianceReportColumnsForSelectedRange(worksheet, "B1", "C1", 25);

            //Replace NULL with blank
            Excel.Range replaceRange = worksheet.Range["J" + Convert.ToString(rowStartIndex), lastColumn + Convert.ToString((rowStartIndex + dtSource.Rows.Count + 1))] as Excel.Range;

            double count = Globals.ThisAddIn.Application.WorksheetFunction.CountIf(replaceRange, "null");
            while (count >= 1)
            {
                removeNulls(replaceRange);
                count = Globals.ThisAddIn.Application.WorksheetFunction.CountIf(replaceRange, "null");
            }

            string[] colRange = varienceList.ToArray();

            Excel.XlFormatConditionOperator rangeConditionOperator = Excel.XlFormatConditionOperator.xlNotEqual;



            if (FAST._varianceFlag == clsInformation.lessThan)
            {
                rangeConditionOperator = Excel.XlFormatConditionOperator.xlLess;
            }
            else if (FAST._varianceFlag == clsInformation.greaterThan)
            {
                rangeConditionOperator = Excel.XlFormatConditionOperator.xlGreater;
            }
            else if (FAST._varianceFlag == clsInformation.both)
            {
                rangeConditionOperator = Excel.XlFormatConditionOperator.xlNotEqual;
            }

            for (int colIndex = 0; colIndex < colRange.Length; colIndex++)
            {
                Excel.Range addPercent = worksheet.Range[colRange[colIndex] + Convert.ToString(varianceReportrowBodyNumber), colRange[colIndex] + rowCnt] as Excel.Range;

                addPercent.NumberFormat = "##0.00\\%";

                ApplyFormatValidations(rowCnt, worksheet, colRange, colIndex, rangeConditionOperator);

            }

            if (FAST._txtProcess == clsInformation.tcpuView)
                // For variance Value Column Comparision
                applyFormatValidationsForVarianceValue(rowCnt, worksheet, dtSource.Columns.Count + 1, Excel.XlFormatConditionOperator.xlGreater);

            for (int i = varianceReportrowBodyNumber; i <= rowCnt; i++)
            {
                if (i % 2 == 1)
                {
                    formatvarienceReportColumns(worksheet, i, lastColumn, "#fff", "#e6faff");
                }
                else
                {
                    formatvarienceReportColumns(worksheet, i, lastColumn, "#fdd49b", "#fff");
                }
                (worksheet.Range["J" + i, lastColumn + i] as Excel.Range).ColumnWidth = 24;

            }

            //Added range,color and row height
            Excel.Range staticRange1 = worksheet.Range["B" + Convert.ToString(varianceReportrowBodyNumber), varianceReportPCLastColumn + Convert.ToString(varianceReportrowBodyNumber)] as Excel.Range;
            staticRange1.Interior.Color = ColorTranslator.FromHtml("#000"); //Black color
            staticRange1.Borders.Color = ColorTranslator.FromHtml("#000");
            staticRange1.RowHeight = "20";

            staticRange1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            staticRange1.VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;

            //Add background color for month column groups
            addColGroupBG(clsInformation.productsRevenueReport, columnCnt, rowStartIndex);
            #endregion

            worksheet.Select(Type.Missing);

            worksheet.Activate();//Activates the current worksheet.

            (worksheet.Range["A1", "A1"] as Excel.Range).ColumnWidth = clsInformation.aColWidth;//Set A1 column width

            (worksheet.Range["B10", "B10"] as Excel.Range).ColumnWidth = clsInformation.bColWidth;//Set B10 column width

            worksheet.Range[clsInformation.constSheetTitleStartRange, clsInformation.constSheetTitleEndRange].Select();

            FAST.updateEvents(true);

            protectSheet(worksheet);

            #endregion

        }



        private static void modifyTCPUVarianceReportMainHeader(Worksheet worksheet, int colValue, int rowStartIndex)
        {



            while (worksheet.Cells[rowStartIndex, colValue].Value != null && !string.IsNullOrEmpty(Convert.ToString(worksheet.Cells[rowStartIndex, colValue].Value)))
            {
                string mnth = Convert.ToString(worksheet.Cells[rowStartIndex, colValue].Value).Replace("-", "'");

                worksheet.Cells[rowStartIndex, colValue].Value = mnth;


                colValue++;

            }

        }

        /// <summary>
        /// Use for Improving the Performance of the Application This method is to apply color format based on condition for selected range
        /// </summary>
        /// <param name="worksheet">Specidfies column name and alternate colors to worksheet</param>
        /// <param name="i"></param>
        /// <param name="lastColumn">Specifies column name</param>
        /// <param name="bColor">Specifies color</param>
        /// <param name="jColor">Specifies color</param>
        private static void formatvarienceReportColumns(Worksheet worksheet, int i, string lastColumn, string bColor, string jColor) //instead of work object we can also use Globals.ThisAddIn.Application
        {
            formatVarianceReportColumnsForSelectedRange(worksheet, "B" + i, varianceReportPCLastColumn + i, ColorTranslator.FromHtml(bColor));
            formatVarianceReportColumnsForSelectedRange(worksheet, varianceReportMonthlyData + i, lastColumn + i, ColorTranslator.FromHtml(jColor));
        }

        /// <summary>
        /// This method is to apply format condition for given range
        /// </summary>
        /// <param name="rowCnt"></param>
        /// <param name="worksheet"></param>
        /// <param name="colRange"></param>
        /// <param name="colIndex"></param>
        /// <param name="formatConditionOperator"></param>
        private static void ApplyFormatValidations(int rowCnt, Worksheet worksheet, string[] colRange, int colIndex, Excel.XlFormatConditionOperator formatConditionOperator)
        {
            Excel.FormatCondition format1 = (Excel.FormatCondition)(worksheet.get_Range(colRange[colIndex] + Convert.ToString(varianceReportrowBodyNumber + 1) + ":" + colRange[colIndex] + rowCnt,
                                            Type.Missing).FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, formatConditionOperator,
                                            Convert.ToDecimal(FAST._variancePercent), Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing));

            format1.Font.Color = ColorTranslator.FromHtml("Red");
        }

        /// <summary>
        /// This mehtod is used to set the color for given range
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="startRage"></param>
        /// <param name="endRange"></param>
        /// <param name="color"></param>
        private static void formatVarianceReportColumnsForSelectedRange(Worksheet worksheet, string startRage, string endRange, Color color)
        {
            Excel.Range reportRange = worksheet.Range[startRage, endRange] as Excel.Range;
            reportRange.Interior.Color = color;
        }
        /// <summary>
        /// This mehtod is used to set the column width for given range
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="startRage"></param>
        /// <param name="endRange"></param>
        /// <param name="columnWidth"></param>
        private static void formatVarianceReportColumnsForSelectedRange(Worksheet worksheet, string startRage, string endRange, int columnWidth)
        {
            Excel.Range reportRange = worksheet.Range[startRage, endRange] as Excel.Range;
            reportRange.ColumnWidth = columnWidth;
        }

        /// <summary>
        /// This mehtod is used to set the font bold, column width, color and text alignment
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="startRage"></param>
        /// <param name="endRange"></param>
        /// <param name="isFontBold"></param>
        /// <param name="columnWidth"></param>
        /// <param name="colorIndex"></param>
        /// <param name="isLeftAlign"></param>
        /// <param name="color"></param>
        private static void formatVarianceReportColumnsForSelectedRange(Worksheet worksheet, string startRage, string endRange, bool isFontBold, int columnWidth, int colorIndex, bool isLeftAlign, Color color)
        {
            Excel.Range reportRange = worksheet.Range[startRage, endRange] as Excel.Range;
            reportRange.Interior.Color = color;
            // For White Text
            reportRange.Font.ColorIndex = colorIndex;
            reportRange.Font.Bold = isFontBold;
            reportRange.ColumnWidth = columnWidth;
            reportRange.Cells.Value = null;

            if (isLeftAlign)
                reportRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
        }

        /// <summary>
        /// This method is used to  set color, text alignment and font bold
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="lastcolumn"></param>
        /// <param name="rowNumber"></param>
        public static void addColGroupBG(string worksheet, int lastcolumn, int rowNumber)
        {
            #region Variance Code

            Worksheet worksheet1 = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[worksheet]);
            if (worksheet1 == null)
                throw new Exception("Worksheet can not be empty when building sheet body");

            int i = varianceReporMonthlyCol, j = 1, monthRow = varianceReportrowBodyNumber - 1, monthCol = varianceReporMonthlyCol + 1;
            string columnNames = Convert.ToString(FAST._dsVarianceReport.Tables[1].Rows[1].ItemArray[0]).Replace("-", "'");

            columnNames = columnNames.Replace("_x0020_", " ");

            fields = columnNames.Split(',');

            while (i < lastcolumn)
            {


                Excel.Range reportRange = worksheet1.Range[worksheet1.Cells[rowNumber, i], worksheet1.Cells[rowNumber, i + 3]] as Excel.Range;
                // Aligining the text for the cells to be in center
                reportRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                // For White Text, 1 for Black ----
                reportRange.Font.ColorIndex = 1;
                reportRange.Font.Bold = true;
                /*gray*/
                reportRange.Interior.Color = (j % 2 != 0) ? ColorTranslator.FromHtml("#fbb450")/*Orange*/ : ColorTranslator.FromHtml("#a6a6a6");

                j += 1;
                i += 4;
            }


            for (int m = 0; m < fields.Length; m++)
            {
                worksheet1.Cells[monthRow, monthCol].Value = fields[m];
                monthCol = monthCol + 4;
            }
            #endregion


        }

        #region Format Conditions for TCPU Variance Value
        /// <summary>
        /// Added this method to apply formating conditions to Previous Scenario and Current Scenario Variance Value
        /// </summary>
        /// <param name="usedRowsCount"></param>
        /// <param name="worksheet"></param>
        /// <param name="lastCoumnCount"></param>
        /// <param name="comparisionOperator"></param>
        private static void applyFormatValidationsForVarianceValue(int usedRowsCount, Worksheet worksheet, int lastCoumnCount, Excel.XlFormatConditionOperator comparisionOperator)
        {
            for (int i = varianceReporMonthlyCol + 2; i < lastCoumnCount; i++)
            {
                string varianceValueColumnName = getColumnName(i);
                //int tcpuVarianceValue = Convert.ToInt32(Convert.ToDecimal(FAST._varianceValue));


                Excel.FormatCondition format1 = (Excel.FormatCondition)(worksheet.get_Range(varianceValueColumnName + Convert.ToString(varianceReportrowBodyNumber + 1) + ":" + varianceValueColumnName + Convert.ToString(varianceReportrowBodyNumber + usedRowsCount),
                                            Type.Missing).FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, comparisionOperator,
                                            Convert.ToDecimal(FAST._varianceValue), Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing));

                format1.Font.Color = ColorTranslator.FromHtml("Red");

                i = i + 3;
            }

        }
        #endregion

        #endregion

        #region Statistics
        public static void buildStatisticsSheetBody(string sheetName, ref ListObject loStatistics, string mode, string txtProcess, DataTable dt,
                                                        ref Chart pie, ref Chart column, ref Chart conecol, string selectedItem, string Scenario)
        {
            Worksheet sheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[sheetName]);

            if (sheet != null)
            {
                unProtect(sheet);
                if (loStatistics != null)
                    sheet.Controls.Remove(loStatistics);

                for (int i = 0; i < sheet.ListObjects.Count; i++)
                    Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[sheet].ListObjects.Item(i + 1).Delete();

                if (pie != null)
                    sheet.Controls.Remove(pie);

                if (column != null)
                    sheet.Controls.Remove(column);

                if (conecol != null)
                    sheet.Controls.Remove(conecol);

                sheet.UsedRange.Clear();

                (sheet.Range["A1", "AZ500"] as Excel.Range).Delete();
                (sheet.Range["A1", "AZ500"] as Excel.Range).Borders.Color = ColorTranslator.FromHtml("#fff");

                buildSheetTitle(sheetName, txtProcess, Scenario, null, null, null, null, null);

                unProtect(sheet);


                string lastcolumn = getColumnName(dt.Columns.Count + 1);

                inputTemplateRowBodyNumber = inputTemplateRowBodyNumber + 25;

                string[] fields = null;
                switch (mode)
                {
                    case clsInformation.statisticsScenarioTag:
                        fields = clsInformation.statisticsScenarioHeader;
                        break;

                    case clsInformation.statisticsInputTypeTag:
                        fields = clsInformation.statisticsInputTypeHeader;
                        break;

                    case clsInformation.statisticsUserTag:
                        fields = clsInformation.statisticsUserHeader;
                        break;
                }

                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    sheet.Cells[inputTemplateRowBodyNumber, i + 2].Value = Convert.ToString(dt.Columns[i].ColumnName);
                }

                int r = inputTemplateRowBodyNumber + 1;
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    for (int k = 0; k < dt.Columns.Count; k++)
                    {
                        sheet.Cells[r, k + 2].Value = Convert.ToString(dt.Rows[i].ItemArray[k]);
                    }
                    r++;
                }

                addStatistics(sheet, dt, ref pie, ref column, ref conecol, selectedItem);
            }


            protectSheet(sheet);

        }
        #endregion

        #region Add Chart Controls and Color Formatting
        private static void addStatistics(Worksheet sheet, DataTable dt, ref Chart pieChartGraph, ref Chart columnChartGraph, ref Chart conicalChartGraph, string selectedItem)
        {

            // Added for Doughnut
            pieChartGraph = sheet.Controls.AddChart(45, 120, 350, 250, "DoughNut");
            pieChartGraph.ChartType = Excel.XlChartType.xlDoughnut;

            //Set chart title
            pieChartGraph.HasTitle = true;
            pieChartGraph.ChartTitle.Text = selectedItem;

            //Add a bar Chart
            columnChartGraph = sheet.Controls.AddChart(405, 120, 350, 250, "column");
            columnChartGraph.ChartType = Excel.XlChartType.xlColumnClustered;

            ////Set chart title
            columnChartGraph.HasTitle = true;
            columnChartGraph.ChartTitle.Text = selectedItem;

            //Add a conical Chart
            conicalChartGraph = sheet.Controls.AddChart(765, 120, 350, 250, "Line");
            conicalChartGraph.ChartType = Excel.XlChartType.xlLineMarkersStacked;

            //Set chart title
            conicalChartGraph.HasTitle = true;
            conicalChartGraph.ChartTitle.Text = selectedItem;

            string columnName = getColumnName(dt.Columns.Count + 1);

            Excel.Range chartRange = null;

            if (selectedItem != clsInformation.statisticsUserLabel)
                chartRange = sheet.get_Range("D" + Convert.ToString(inputTemplateRowBodyNumber), columnName + Convert.ToString(inputTemplateRowBodyNumber + dt.Rows.Count));
            else
                chartRange = sheet.get_Range("E" + Convert.ToString(inputTemplateRowBodyNumber), columnName + Convert.ToString(inputTemplateRowBodyNumber + dt.Rows.Count));

            pieChartGraph.SetSourceData(chartRange, Type.Missing);
            columnChartGraph.SetSourceData(chartRange, Type.Missing);
            conicalChartGraph.SetSourceData(chartRange, Type.Missing);

            applyFormat(sheet, dt, 3, columnName);

            sheet.Select();

            (sheet.Range["A1", "A1"] as Excel.Range).ColumnWidth = 5;

            sheet.Range["B2", "B2"].Select();
        }

        private static void applyFormat(Worksheet Sheet, DataTable dt, int lockCount, string columnValue)
        {
            Excel.Range applyHeaderColor = Sheet.get_Range("B" + Convert.ToString(inputTemplateRowBodyNumber), columnValue + Convert.ToString(inputTemplateRowBodyNumber)) as Excel.Range;
            applyHeaderColor.Font.Color = ColorTranslator.FromHtml("#fff");
            applyHeaderColor.Interior.Color = ColorTranslator.FromHtml("#000");
            applyHeaderColor.RowHeight = 20;
            applyHeaderColor.ColumnWidth = 22;

            string lastColumn = getColumnName(lockCount);

            Excel.Range rng1 = null, rng2 = null;

            for (int i = inputTemplateRowBodyNumber + 1; i <= inputTemplateRowBodyNumber + dt.Rows.Count; i++)
            {
                rng1 = Sheet.get_Range("B" + Convert.ToString(i), lastColumn + Convert.ToString(i)) as Excel.Range;

                lastColumn = getColumnName(lockCount + 1);

                rng2 = Sheet.get_Range(lastColumn + Convert.ToString(i), columnValue + Convert.ToString(i)) as Excel.Range;


                if (i % 2 == 0)
                {
                    rng1.Interior.Color = ColorTranslator.FromHtml("#fdd49b");
                    rng2.Interior.ColorIndex = 19;
                }
                else
                {
                    rng1.Interior.Color = ColorTranslator.FromHtml("#fdfdfd");
                    rng2.Interior.ColorIndex = 19;
                }

            }
        }

        #endregion

        public static void prepareTestSheet(string shtName, ref ListObject lo, DataTable dtSource)
        {
            Worksheet worksheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[shtName]);
            if (worksheet == null)
                throw new Exception("Worksheet can not be empty when building sheet body");

            unProtect(worksheet);

            if (lo != null)
                worksheet.Controls.Remove(lo);

            for (int i = 0; i < Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[shtName].ListObjects.Count; i++)
                Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[shtName].ListObjects.Item(i + 1).Delete();


            // Clearing the Used Range
            worksheet.UsedRange.Clear();

            unProtect(worksheet);

            FAST.updateEvents(false);

            int rowStartIndex = 5;

            FAST._lastColumnName = getColumnName(dtSource.Columns.Count);

            // For Pivot Data
            Excel.Range lstObjRange = worksheet.Range["B" + Convert.ToString(rowStartIndex), (FAST._lastColumnName + Convert.ToString(rowStartIndex + dtSource.Rows.Count))] as Excel.Range;
            lstObjRange.Clear();

            string listObjectName = "ListObject" + shtName;

            lo = worksheet.Controls.AddListObject(lstObjRange, listObjectName);

            string columnNames = null;

            columnNames = Convert.ToString(FAST._dsDownloadData.Tables[0].Rows[0].ItemArray[0]).Replace(" ", string.Empty);

            string[] fields = columnNames.Split(',');

            for (int i = 0; i < lo.ListColumns.Count; i++)
            {
                lo.ListColumns[i + 1].Name = fields[i];
            }

            BindingSource bs = new BindingSource();
            bs.DataSource = dtSource;
            lo.SetDataBinding(bs, "", fields);

            // This  Method will work only for Accounting View
            //if (FAST._txtProcess == "Accounting View")
            //{
            //    modifyHeaders(worksheet);
            //}
            //else
            //{
            //    modifyTCPUHeaders(worksheet);
            //}

            // Adding the Formula Name to the respected Column.
            //worksheet.Cells[inputTemplateRowBodyNumber, formulaColumnName].Value = clsInformation.formulaColumnName;


            //setSlideDeckListObjectStyle(worksheet, rowStartIndex + 1, dtSource.Rows.Count, FAST._lastColumnName);

            //double count = Globals.ThisAddIn.Application.WorksheetFunction.CountIf(lstObjRange, "null");

            //while (count >= 1)
            //{
            //    removeNulls(lstObjRange);
            //    count = Globals.ThisAddIn.Application.WorksheetFunction.CountIf(lstObjRange, "null");
            //}

            //Excel.Range verifyNullsRange = worksheet.Range[formulaNextColumn + Convert.ToString(rowStartIndex + 1), (FAST._lastColumnName + Convert.ToString(rowStartIndex + dtSource.Rows.Count))] as Excel.Range;
            //count = Globals.ThisAddIn.Application.WorksheetFunction.CountIf(verifyNullsRange, "null");

            //if (count >= 1)
            //{
            //    removeInputTemplateNulls(dtSource, worksheet);
            //}

            //lockCorrespondCells(worksheet, dtSource, FAST._lastColumnName);
            FAST.updateEvents(true);

            //formatCells(worksheet, dtSource.Rows.Count, FAST._lastColumnName);

            //if (FAST._txtProcess == clsInformation.tcpuView)
            //{
            //    lockAndSetLifeTimeValueColumnForMSRP(worksheet.Name);
            //}
            Excel.Range unlock = worksheet.Range["F" + Convert.ToString(rowStartIndex + 1), ("F" + Convert.ToString(rowStartIndex + dtSource.Rows.Count))] as Excel.Range;
            unlock.Locked = false;
            unlock.FormulaHidden = false;

            Globals.ThisAddIn.Application.ActiveWindow.DisplayFormulas = true;

            FAST.updateEvents(false);

            worksheet.Activate();

            Globals.ThisAddIn.Application.AutoCorrect.AutoFillFormulasInLists = false;

            worksheet.AutoFilterMode = false;

            (worksheet.Range["A1", "A1"] as Excel.Range).ColumnWidth = 5;

            worksheet.Range["B2", "B2"].Select();

            protectSheet(worksheet);
        }

        public static void bindDataUsingObjectArray(string shtName, DataTable dt)
        {
            Worksheet worksheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[shtName]);
            if (worksheet == null)
                throw new Exception("Worksheet can not be empty when building sheet body");

            object[,] arr = new object[dt.Rows.Count, dt.Columns.Count];
            for (int r = 0; r < dt.Rows.Count; r++)
            {
                DataRow dr = dt.Rows[r];
                for (int c = 0; c < dt.Columns.Count; c++)
                {
                    arr[r, c] = dr[c];
                }
            }

            Excel.Range c1 = (Excel.Range)worksheet.Cells[2, 1];
            Excel.Range c2 = (Excel.Range)worksheet.Cells[2 + dt.Rows.Count - 1, dt.Columns.Count];
            Excel.Range range = worksheet.get_Range(c1, c2);

            range.Value = arr;
        }
    }

}



