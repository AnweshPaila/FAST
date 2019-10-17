using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelTool = Microsoft.Office.Tools.Excel;
using System.Data;
using System.Windows.Forms;
using System.Drawing;
using Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.IO;


namespace Test_WorkBookOpen.Classes
{
    public class ClsPromotions
    {
        #region Global variables
        public static DataTable dtPromoMaster, dtPromoRows, dtVdpRows, dtTcpuRows, dtPromoFormulas, dtBransonPromotionsRows;
        const string promoMasterTableName = "PromoTool", promoRowsTableName = "PromoRows", vdpRowsTableNAme = "VDPRows", tcpuRowsTableName = "TCPURows", promoFormulaTableName = "PromoFormulas", BransonPromotionsRows = "BransonPromotionsRows";
        static string vdpColHeaders, tcpuColHeaders, promoColHeaders, dropDownData_ProgramList, dropDownData_TCPUCodeNameList, dropDownData_ChannelList, dropDownData_PromoTypeList;
        static DataSet dsDeviceTypes = null, dsPromotions = null;
        ListObject lstObjectTcpu = null, lstObjectVdp = null, lo = null, lstObjectBransonPromotions = null;

        public static string tcpuCodeNameColumn; // Added by Nihar on 11/30/2017
        public static long promoBodyEndRows = 0;
        public static int tcpuCodeNameListTableId = 0;
        #endregion

        #region Promotions View
        // Added by Nihar
        public void promotionsView(string aliasId, string viewId, string countryId, string deviceTypeId, string view, string countryLabel, string deviceTypeLabel)
        {
            try
            {
                FAST.updateEvents(false);

                FAST._IsPromotionErrorHit = false;

                dsPromotions = new DataSet();


                //string jsonResponse = 
                clsDataSheet.WriteToLogFile(" calling Service : " + DateTime.Now.ToString());

                //string jsonResponse = 

                dsPromotions = FASTWebServiceAdapter.getDownloadDataForPromotions(aliasId, viewId, countryId, deviceTypeId, view);

                if (dsPromotions == null)
                    return;

                if (dsPromotions.Tables.Count > 0)
                {
                    FAST._promoCountryValue = countryId;
                    FAST._promoDeviceTypeValue = deviceTypeId;
                    FAST._promoDownloadCountryValueForOfflineOnline = FAST._promoCountryValue;
                    FAST._promoDownloadDeviceTypeForOfflineOnline = FAST._promoDeviceTypeValue;


                    // Get all column Headers from DataTable
                    dtPromoMaster = dsPromotions.Tables[promoMasterTableName];
                    dtPromoRows = dsPromotions.Tables[promoRowsTableName];
                    dtVdpRows = dsPromotions.Tables[vdpRowsTableNAme];
                    dtTcpuRows = dsPromotions.Tables[tcpuRowsTableName];
                    dtPromoFormulas = dsPromotions.Tables[promoFormulaTableName];
                    //Added by Praveen
                    dtBransonPromotionsRows = dsPromotions.Tables[BransonPromotionsRows];

                    int programListNumber = 0, tcpuCodeNameListnumber = 0;

                    for (int i = 0; i < dsPromotions.Tables.Count; i++)
                    {
                        if (dsPromotions.Tables[i].TableName == "ProgramList")
                            programListNumber = i;
                        if (dsPromotions.Tables[i].TableName == "TCPUCodeNameList")
                            tcpuCodeNameListnumber = i;
                    }

                    if (dtPromoMaster.Rows.Count > 0)
                    {
                        vdpColHeaders = dtPromoMaster.Rows[0]["VDPHeader"].ToString().ReplaceNewLine();
                        tcpuColHeaders = dtPromoMaster.Rows[0]["TCPUHeader"].ToString().ReplaceNewLine();
                        promoColHeaders = dtPromoMaster.Rows[0]["PromoHeader"].ToString().ReplaceNewLine();

                        dropDownData_ProgramList = null;

                        // iterating for Program List and also For TcpuCodeName
                        for (int i = 0; i < dsPromotions.Tables[programListNumber].Rows.Count; i++)
                        {
                            for (int j = 0; j < dsPromotions.Tables[programListNumber].Columns.Count - 1; j++)
                            {
                                string value = Convert.ToString(dsPromotions.Tables[programListNumber].Rows[i].ItemArray[j]);

                                if (value != "" && value != null)
                                    dropDownData_ProgramList += "<ProgramList>" + dsPromotions.Tables[programListNumber].Columns[j].ColumnName + "|" + value + "</ProgramList>";
                            }
                        }
                        dropDownData_TCPUCodeNameList = null;

                        for (int i = 0; i < dsPromotions.Tables[tcpuCodeNameListnumber].Rows.Count; i++)
                        {
                            for (int j = 0; j < dsPromotions.Tables[tcpuCodeNameListnumber].Columns.Count - 1; j++)
                            {
                                string value = Convert.ToString(dsPromotions.Tables[tcpuCodeNameListnumber].Rows[i].ItemArray[j]);

                                if (value != "" && value != null)
                                    dropDownData_TCPUCodeNameList += "<TcpuCodeNameList>" + dsPromotions.Tables[tcpuCodeNameListnumber].Columns[j].ColumnName + "|" + value + "</TcpuCodeNameList>";
                            }
                        }

                        dropDownData_ChannelList = Convert.ToString(dtPromoMaster.Rows[0]["ChannelList"]);
                        dropDownData_PromoTypeList = Convert.ToString(dtPromoMaster.Rows[0]["PromoTypeList"]);

                        dropDownData_ProgramList = "<DeviceTypes><Program>" + dropDownData_ProgramList +
                                                    "</Program><TcpuCodeName>" + dropDownData_TCPUCodeNameList +
                                                     "</TcpuCodeName><Channel><ChannelType>" + dropDownData_ChannelList +
                                                     "</ChannelType></Channel><Promo><PromoType>" + dropDownData_PromoTypeList +
                                                     "</PromoType></Promo></DeviceTypes>";

                        dsDeviceTypes = new DataSet();
                        using (StringReader stringReader = new StringReader(dropDownData_ProgramList))
                        {
                            dsDeviceTypes.ReadXml(stringReader);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Required Data is missing....");
                        return;
                    }
                    ////Added by Praveen
                    //clsDataSheet.WriteToLogFile(" calling BransonPromotionsSheet : " + DateTime.Now.ToString());

                    ExcelTool.Workbook excelWorkbook = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook);

                    List<string> sheetNames = new List<string>();
                    foreach (Excel.Worksheet sheet in excelWorkbook.Sheets)
                    {
                        sheetNames.Add(sheet.Name);
                    }


                    //if (sheetNames.Contains(clsInformation.PROMO_INPUT_TOOL))
                    //    generatePromoInputTemplateSheet(countryLabel, deviceTypeLabel, aliasId);
                    //if (sheetNames.Contains(clsInformation.TCPU))
                    //    generateTcpuSheet(countryLabel, deviceTypeLabel, aliasId);
                    //if (sheetNames.Contains(clsInformation.VDP))
                    //    generateVdpSheet(countryLabel, deviceTypeLabel, aliasId);
                    //if (sheetNames.Contains(clsInformation.bransonPromotions))
                    //    generateBransonPromotionsSheet(countryLabel, deviceTypeLabel, aliasId);

                    if (sheetNames.Contains(clsInformation.bransonPromotions))
                        generateBransonPromotionsSheet(countryLabel, deviceTypeLabel, aliasId);
                    if (sheetNames.Contains(clsInformation.VDP))
                        generateVdpSheet(countryLabel, deviceTypeLabel, aliasId);
                    if (sheetNames.Contains(clsInformation.TCPU))
                        generateTcpuSheet(countryLabel, deviceTypeLabel, aliasId);
                    if (sheetNames.Contains(clsInformation.PROMO_INPUT_TOOL))
                        generatePromoInputTemplateSheet(countryLabel, deviceTypeLabel, aliasId);

                }

            }
            catch (Exception ex)
            {
                FAST.errorLog(ex.Message, "Promotion_Planning_Excel_promotionsView");
                FAST.handleAlerts(ex.Message);
                FAST._IsPromotionErrorHit = true;
            }
            finally
            {
                FAST.updateEvents(true);
            }
        }
        #endregion

        #region Refresh Vdp Tcpu
        public void refreshVdpTcpu(string aliasId, string viewId, string countryId, string deviceTypeId, string view, string countryLabel, string deviceTypeLabel)
        {
            try
            {
                FAST.updateEvents(false);

                DataSet dsVdpTcpu = new DataSet();

                dsVdpTcpu = FASTWebServiceAdapter.refreshVdpTcpuForPromoInputTool(aliasId, viewId, countryId, deviceTypeId, view);


                if (dsVdpTcpu.Tables.Count > 1)
                {
                    dtVdpRows = dsVdpTcpu.Tables[vdpRowsTableNAme];
                    dtTcpuRows = dsVdpTcpu.Tables[tcpuRowsTableName];
                }
                if (!dsVdpTcpu.Tables.Contains(vdpRowsTableNAme))
                {
                    dtVdpRows = null;
                }
                if (!dsVdpTcpu.Tables.Contains(tcpuRowsTableName))
                {
                    dtTcpuRows = null;
                }


                if (dsVdpTcpu.Tables[0].Rows.Count > 0)
                {

                    vdpColHeaders = dsVdpTcpu.Tables[0].Rows[0]["VDPHeader"].ToString().ReplaceNewLine();
                    tcpuColHeaders = dsVdpTcpu.Tables[0].Rows[0]["TCPUHeader"].ToString().ReplaceNewLine();
                }
                else
                {

                    MessageBox.Show("Required Data is missing....");
                    return;
                }

                generateVdpSheet(countryLabel, deviceTypeLabel, aliasId);
                generateTcpuSheet(countryLabel, deviceTypeLabel, aliasId);

                // Selecting the PromoTool Sheet Once the work is done
                //Worksheet worksheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[clsInformation.PROMO_INPUT_TOOL]);

                //if (worksheet != null)
                //    worksheet.Select();

            }
            catch (Exception ex)
            {
                FAST.errorLog(ex.Message, "Promotion_Planning_Excel_promotionsView");
                FAST.handleAlerts(ex.Message);
                FAST._IsPromotionErrorHit = true;
            }
            finally
            {
                FAST.updateEvents(true);
            }

        }
        #endregion

        #region Generate VdpbSheet
        private void generateVdpSheet(string country, string deviceType, string aliasId)
        {
            try
            {

                FAST.updateEvents(false);

                Worksheet worksheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[clsInformation.VDP]);


                if (worksheet == null)
                    throw new Exception("Worksheet can not be empty when building sheet body");

                if (worksheet.Visible != Excel.XlSheetVisibility.xlSheetVeryHidden)
                    worksheet.Select();


                if (dtVdpRows == null)
                {
                    unProtect(worksheet);
                    if (lstObjectVdp != null)
                        worksheet.Controls.Remove(lstObjectVdp);

                    if (worksheet.UsedRange != null)
                        worksheet.UsedRange.Clear();

                    protectSheet(worksheet);

                    if (worksheet.Visible != Excel.XlSheetVisibility.xlSheetVeryHidden)
                        FAST.displayAlerts("No Data to generate VDP Sheet", 1);

                    return;
                }

                unProtect(worksheet);

                bindDataToSheetUsingListObject(ref lstObjectVdp, vdpColHeaders, "11", dtVdpRows, worksheet);

                //Added by Praveen
                int startRow = 12;
                int endRow = dtVdpRows.Rows.Count;

                //Converting date format to general for vdp_BaseLine
                Excel.Range range = worksheet.get_Range("N" + startRow, "N" + (startRow + endRow)) as Excel.Range;
                range.NumberFormat = "General";

                string lastColumnForPromoVdpSheet = clsManageSheet.getColumnName(dtVdpRows.Columns.Count + 1);

                sheetTitleSummary(worksheet, generateFilterSummary(country, deviceType, "local", aliasId));

                resettingColumnWidth(worksheet, 2, 11);

                if (worksheet.Visible != Excel.XlSheetVisibility.xlSheetVeryHidden)
                    applyingBackGroundColorToSheet(worksheet, 11, 11 + dtVdpRows.Rows.Count, "B", clsManageSheet.getColumnName(dtVdpRows.Columns.Count));

                //Hide pcid col here
                Excel.Range hideRange = worksheet.get_Range(lastColumnForPromoVdpSheet + Convert.ToString("1"), lastColumnForPromoVdpSheet + Convert.ToString("1")) as Excel.Range;
                hideRange.EntireColumn.Hidden = true;

                splitAndFreezPanes(worksheet, 11);

                (worksheet.Range["A1", "A1"] as Excel.Range).ColumnWidth = 5;

                if (worksheet.Visible != Excel.XlSheetVisibility.xlSheetVeryHidden)
                    worksheet.Range["B2", "B2"].Select();

                protectSheet(worksheet);
            }
            catch (Exception ex)
            {
                FAST.errorLog(ex.Message, "Promotion_Planning_Excel_promotionsView");
                FAST.handleAlerts(ex.Message);
                FAST._IsPromotionErrorHit = true;
            }
            finally
            {
                FAST.updateEvents(true);
            }
        }
        #endregion

        #region Generate Tcpu Sheet
        private void generateTcpuSheet(string country, string deviceType, string aliasId)
        {
            try
            {
                FAST.updateEvents(false);


                Worksheet worksheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[clsInformation.TCPU]);

                if (worksheet == null)
                    throw new Exception("Worksheet can not be empty when building sheet body");

                if (worksheet.Visible != Excel.XlSheetVisibility.xlSheetVeryHidden)
                    worksheet.Select();

                if (dtTcpuRows == null)
                {
                    unProtect(worksheet);
                    if (lstObjectTcpu != null)
                        worksheet.Controls.Remove(lstObjectTcpu);

                    if (worksheet.UsedRange != null)
                        worksheet.UsedRange.Clear();

                    protectSheet(worksheet);

                    if (worksheet.Visible != Excel.XlSheetVisibility.xlSheetVeryHidden)
                        FAST.displayAlerts("No Data to generate TCPU Sheet", 1);

                    return;
                }

                unProtect(worksheet);

                bindDataToSheetUsingListObject(ref lstObjectTcpu, tcpuColHeaders, "11", dtTcpuRows, worksheet);

                string lastColumnForPromoTcpuSheet = clsManageSheet.getColumnName(dtTcpuRows.Columns.Count + 1);

                sheetTitleSummary(worksheet, generateFilterSummary(country, deviceType, "local", aliasId));

                resettingColumnWidth(worksheet, 2, 11);

                if (worksheet.Visible != Excel.XlSheetVisibility.xlSheetVeryHidden)
                    applyingBackGroundColorToSheet(worksheet, 11, 11 + dtTcpuRows.Rows.Count, "B", clsManageSheet.getColumnName(dtTcpuRows.Columns.Count));

                //Hide pcid col here
                Excel.Range hideRange = worksheet.get_Range(lastColumnForPromoTcpuSheet + Convert.ToString("1"), lastColumnForPromoTcpuSheet + Convert.ToString("1")) as Excel.Range;
                hideRange.EntireColumn.Hidden = true;

                splitAndFreezPanes(worksheet, 11);

                (worksheet.Range["A1", "A1"] as Excel.Range).ColumnWidth = 5;

                if (worksheet.Visible != Excel.XlSheetVisibility.xlSheetVeryHidden)
                    worksheet.Range["B2", "B2"].Select();

                protectSheet(worksheet);
            }
            catch (Exception ex)
            {
                FAST.errorLog(ex.Message, "Promotion_Planning_Excel_promotionsView");
                FAST.handleAlerts(ex.Message);
                FAST._IsPromotionErrorHit = true;
            }
            finally
            {
                FAST.updateEvents(true);
            }
        }
        #endregion

        #region Generate Bransonsheet
        private void generateBransonPromotionsSheet(string country, string deviceType, string aliasId)
        {
            try
            {
                FAST.updateEvents(false);

                Worksheet worksheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[clsInformation.bransonPromotions]);


                if (worksheet == null)
                    throw new Exception("Worksheet can not be empty when building sheet body");

                if (worksheet.Visible != Excel.XlSheetVisibility.xlSheetVeryHidden)
                    worksheet.Select();


                if (dtBransonPromotionsRows == null)
                {
                    unProtect(worksheet);
                    if (lstObjectBransonPromotions != null)
                        worksheet.Controls.Remove(lstObjectBransonPromotions);

                    if (worksheet.UsedRange != null)
                        worksheet.UsedRange.Clear();

                    protectSheet(worksheet);

                    if (worksheet.Visible != Excel.XlSheetVisibility.xlSheetVeryHidden)
                        FAST.displayAlerts("No Data to generate BransonPromotions Sheet", 1);

                    return;
                }

                unProtect(worksheet);

                //TODO : use this code untill we get headers from service layer
                // hare coded because of format
                string BransonPromotionsColHeaders = clsInformation.bransonHeader;

                //string[] BransonPromotionsColHeaders = dtBransonPromotionsRows.Columns.Cast<DataColumn>().
                //                               Select(column => column.ColumnName).
                //                               ToArray();

                // var BransonColHeaders = string.Join(",", BransonPromotionsColHeaders);

                bindDataToSheetUsingListObject(ref lstObjectBransonPromotions, BransonPromotionsColHeaders, "11", dtBransonPromotionsRows, worksheet);

                int rowStartIndex = 12;

                int lastRow = dtBransonPromotionsRows.Rows.Count;

                Excel.Range lstObjRange = worksheet.Range[clsInformation.constStartRange + Convert.ToString(rowStartIndex), (clsInformation.constStartRange + Convert.ToString(rowStartIndex + lastRow))] as Excel.Range;

                lstObjRange.NumberFormat = clsInformation.constDateFormat;


                string lastColumnForPromoBrandsonSheet = clsManageSheet.getColumnName(dtBransonPromotionsRows.Columns.Count + 1);

                sheetTitleSummary(worksheet, generateFilterSummary(country, deviceType, "local", aliasId));

                resettingColumnWidth(worksheet, 2, 11);

                if (worksheet.Visible != Excel.XlSheetVisibility.xlSheetVeryHidden)
                    applyingBackGroundColorToSheet(worksheet, 11, 11 + dtBransonPromotionsRows.Rows.Count, "B", clsManageSheet.getColumnName(dtBransonPromotionsRows.Columns.Count));

                //Hide pcid col here
                Excel.Range hideRange = worksheet.get_Range(lastColumnForPromoBrandsonSheet + Convert.ToString("1"), lastColumnForPromoBrandsonSheet + Convert.ToString("1")) as Excel.Range;
                hideRange.EntireColumn.Hidden = true;

                splitAndFreezPanes(worksheet, 11);

                (worksheet.Range["A1", "A1"] as Excel.Range).ColumnWidth = 5;

                if (worksheet.Visible != Excel.XlSheetVisibility.xlSheetVeryHidden)
                    worksheet.Range["B2", "B2"].Select();

                protectSheet(worksheet);
            }
            catch (Exception ex)
            {
                FAST.errorLog(ex.Message, "Promotion_Planning_Excel_generateBransonPromotionsSheet");
                FAST.handleAlerts(ex.Message);
                FAST._IsPromotionErrorHit = true;
            }
            finally
            {
                FAST.updateEvents(true);
            }
        }

        #endregion

        #region Generate Promo Input Template Sheet
        private void generatePromoInputTemplateSheet(string country, string deviceType, string aliasId)
        {
            try
            {
                FAST.updateEvents(false);


                Worksheet worksheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[clsInformation.PROMO_INPUT_TOOL]);

                if (worksheet == null)
                    throw new Exception("Worksheet can not be empty when building sheet body");

                if (worksheet.Visible != Excel.XlSheetVisibility.xlSheetVeryHidden)
                    worksheet.Select();

                unProtect(worksheet);

                dtPromoRows = resetDataTableColumnsOrder(dtPromoRows, promoColHeaders);

                string lastColumnForPromo = clsManageSheet.getColumnName(dtPromoRows.Columns.Count);
                //  clsDataSheet.WriteToLogFile("START-------------------------------------------- ");
                //  clsDataSheet.WriteToLogFile("bindingDataToSheetUsingIteration : " + DateTime.Now.ToString());

                bindingDataToSheetUsingIteration(2, 11, dtPromoRows, worksheet, ref lo);

                // clsDataSheet.WriteToLogFile("sheetTitleSummary : " + DateTime.Now.ToString());
                sheetTitleSummary(worksheet, generateFilterSummary(country, deviceType, "local", aliasId));

                resettingColumnWidth(worksheet, 2, 11);

                // clsDataSheet.WriteToLogFile("applyingBackGroundColorToSheet : " + DateTime.Now.ToString());
                if (worksheet.Visible != Excel.XlSheetVisibility.xlSheetVeryHidden)
                    applyingBackGroundColorToSheet(worksheet, 11, 11 + dtPromoRows.Rows.Count, "B", lastColumnForPromo);

                //  clsDataSheet.WriteToLogFile("dataValidationsForPromoInputTemplate : " + DateTime.Now.ToString());
                dataValidationsForPromoInputTemplate(worksheet, 11, 11 + dtPromoRows.Rows.Count);

                // clsDataSheet.WriteToLogFile("applyingFormulasToPromoInputToolSheet : " + DateTime.Now.ToString());


                // applyingFormulasToPromoInputToolSheet(worksheet, dtPromoRows, 11, 2);

                //anwesh 24/08/2019
                ExcelTool.Workbook excelWorkbook = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook);

                List<string> sheetNames = new List<string>();
                foreach (Excel.Worksheet sheet in excelWorkbook.Sheets)
                {
                    sheetNames.Add(sheet.Name);
                }

                applyingFormulasToPromoInputToolSheet(worksheet, dtPromoRows, 11, 2);


                addDropdownListToColumns(worksheet, 2, 12, 11 + dtPromoRows.Rows.Count, "C", "E", "F", "G", "H");
                makingColumnsReadOnly(worksheet, dtPromoRows, dtPromoFormulas, 11, 11 + dtPromoRows.Rows.Count);

                promoBodyEndRows = 11 + dtPromoRows.Rows.Count;

                // Added by Nihar for TCPU Code Name  Column
                tcpuCodeNameColumn = "F";

                // clsDataSheet.WriteToLogFile("makingColumnsReadOnly : " + DateTime.Now.ToString());

                //Hide pcid col here
                Excel.Range hideRange = worksheet.get_Range(lastColumnForPromo + Convert.ToString("1"), lastColumnForPromo + Convert.ToString("1")) as Excel.Range;
                hideRange.EntireColumn.Hidden = true;

                splitAndFreezPanes(worksheet, 11);

                (worksheet.Range["A1", "A1"] as Excel.Range).ColumnWidth = 5;

                if (worksheet.Visible != Excel.XlSheetVisibility.xlSheetVeryHidden)
                    worksheet.Range["B2", "B2"].Select();

                Excel.Range rg = (Excel.Range)worksheet.Cells[11, "J"];
                rg.EntireColumn.NumberFormat = "MM/DD/YYYY";

                Excel.Range rg1 = (Excel.Range)worksheet.Cells[11, "K"];
                rg1.EntireColumn.NumberFormat = "MM/DD/YYYY";

                protectSheet(worksheet);
            }
            catch (Exception ex)
            {
                FAST.errorLog(ex.Message, "Promotion_Planning_Excel_promotionsView");
                FAST.handleAlerts(ex.Message);
                FAST._IsPromotionErrorHit = true;
            }
            finally
            {
                FAST.updateEvents(true);
            }

            // clsDataSheet.WriteToLogFile("-----------------------------------END : " + DateTime.Now.ToString());
        }
        #endregion

        #region Split And Freez Panes
        private void splitAndFreezPanes(Worksheet worksheet, int freezRow)
        {
            worksheet.Application.ActiveWindow.SplitRow = freezRow;
            worksheet.Application.ActiveWindow.FreezePanes = true;
            worksheet.Application.ActiveWindow.Zoom = 80;

        }
        #endregion

        #region Generate Filter Summary
        private Dictionary<string, string> generateFilterSummary(string country, string deviceType, string currency, string aliasId)
        {
            Dictionary<string, string> filterSummary = new Dictionary<string, string>();

            filterSummary.Add("Country", country);
            filterSummary.Add("Device Type", deviceType);
            filterSummary.Add("User", aliasId);
            filterSummary.Add("Download DateTime:", DateTime.Now.ToString());
            //filterSummary.Add("Currency", currency);

            return filterSummary;


        }
        #endregion

        #region Sheet Title Summary
        private void sheetTitleSummary(Worksheet sheet, Dictionary<string, string> filterSummary)
        {
            Excel.Range sheetTitleRange = sheet.Range[clsInformation.constSheetTitleStartRange, clsInformation.constSheetTitleEndRange] as Excel.Range;
            sheetTitleRange.Merge(false);                                        // Here we are not Merging the cells for Sheet Title
            sheetTitleRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; // Aligining the text for the cells to be in center
            sheetTitleRange.VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;  // Aligining the text for the cells to be in center
            //sheetTitleRange.ColumnWidth = 24;
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

            int rowNumber = 5, columnNumber = 2;

            foreach (var item in filterSummary)
            {
                sheet.Cells[rowNumber, columnNumber].Value = item.Key;
                sheet.Cells[rowNumber, columnNumber].Font.Bold = true;
                sheet.Cells[rowNumber, columnNumber + 1].Value = item.Value;
                sheet.Cells[rowNumber, columnNumber + 1].HorizontalAlignment = XlHAlign.xlHAlignLeft;

                rowNumber++;
            }
        }
        #endregion

        #region bindData To Sheet Using List Object
        private void bindDataToSheetUsingListObject(ref ListObject controlListObject, string header, string startRow, DataTable dt, Worksheet worksheet)
        {
            if (controlListObject != null)
                worksheet.Controls.Remove(controlListObject);

            if (worksheet.UsedRange != null)
                worksheet.UsedRange.Clear();

            string[] fields = header.Split(',');
            string lastColumn = clsManageSheet.getColumnName(fields.Length + 1);
            Excel.Range lstObjRange = worksheet.Range["B" + startRow, (lastColumn + (Convert.ToInt32(startRow) + dt.Rows.Count))] as Excel.Range;
            lstObjRange.Clear();
            worksheet.UsedRange.Clear();
            string listObjectName = "ListObject_" + worksheet.Name;
            // controlListObject = worksheet.Controls.AddListObject(lstObjRange, listObjectName);

            controlListObject = worksheet.Controls.AddListObject(lstObjRange, listObjectName);
            controlListObject.DataSource = dt;
            controlListObject.AutoSetDataBoundColumnHeaders = true;

            //if (listObjectName != "ListObject_BransonPromotions")
            //{
            BindingSource bs = new BindingSource();
            bs.DataSource = dt;
            controlListObject.SetDataBinding(bs, "", fields);
            //}



            //added By Praveen
            lstObjRange.HorizontalAlignment = XlHAlign.xlHAlignLeft;

            // commented by 
            /*  for (int i = 0; i < controlListObject.ListColumns.Count; i++)
                {
                    controlListObject.ListColumns[i + 1].Name = fields[i];
                }

                
                lstObjRange.Columns.AutoFit();
                */


        }
        #endregion

        #region Binding Data To Sheet Using Iteration

        private void bindingDataToSheetUsingIteration(int startColumn, int startRow, DataTable dt, Worksheet worksheet, ref ExcelTool.ListObject lo)
        {
            if (lo != null)
                worksheet.Controls.Remove(lo);

            if (worksheet.UsedRange != null)
                worksheet.UsedRange.Clear();

            promoColHeaders = promoColHeaders.Replace(" ", "");

            string[] colHeaderList = promoColHeaders.Split(',');

            string lastColumnName = clsManageSheet.getColumnName(colHeaderList.Count() + 1);

            int rowNumber = startRow, columnNumber = startColumn;

            for (int i = 0; i < dt.Columns.Count - 1; i++)
            {
                worksheet.Cells[startRow, startColumn].Value = dt.Columns[i].ColumnName;
                startColumn++;
            }
            startRow = rowNumber + 1;

            //Excel.Range bodyRange1 = worksheet.get_Range("B11", "AP111") as Excel.Range;

            Excel.Range bodyRange1 = worksheet.get_Range("B" + Convert.ToString(rowNumber), clsManageSheet.getColumnName(dt.Columns.Count) + Convert.ToString(11 + dtPromoRows.Rows.Count)) as Excel.Range;

            //Excel.Range bodyRange1 = worksheet.get_Range("B"+ startRow, lastColumnName + dtPromoRows.Rows.Count) as Excel.Range;

            string listObjectName = "ListObject_" + worksheet.Name.ReplaceEmptyWithUnderScore();


            lo = worksheet.Controls.AddListObject(bodyRange1, listObjectName);

            BindingSource bs = new BindingSource();
            bs.DataSource = dt;
            lo.SetDataBinding(bs, "", colHeaderList);

            Globals.ThisAddIn.Application.AutoCorrect.AutoFillFormulasInLists = false;


            bodyRange1.HorizontalAlignment = XlHAlign.xlHAlignLeft;

            #region Commented Code
            //object[,] arr = new object[dt.Rows.Count, dt.Columns.Count];
            //for (int r = 0; r < dt.Rows.Count; r++)
            //{
            //    DataRow dr = dt.Rows[r];
            //    for (int c = 0; c < dt.Columns.Count; c++)
            //    {
            //        arr[r, c] = dr[c];
            //    }
            //}

            //Excel.Range c1 = (Excel.Range)worksheet.Cells[rowNumber+1, 2];
            //Excel.Range c2 = (Excel.Range)worksheet.Cells[rowNumber+1 + dt.Rows.Count - 1, dt.Columns.Count];
            //Excel.Range range = worksheet.get_Range(c1, c2);

            //range.Value = arr;

            //for (int i = 0; i < dt.Columns.Count - 1; i++)
            //{
            //    worksheet.Cells[startRow, startColumn].Value = dt.Columns[i].ColumnName;
            //    startColumn++;
            //}




            //startRow = rowNumber; startColumn = columnNumber;
            //string lastColumn = clsManageSheet.getColumnName(dt.Columns.Count - 1);
            //Excel.Range bodyRange = worksheet.get_Range("B" + Convert.ToString(rowNumber), lastColumn + Convert.ToString(11 + dtPromoRows.Rows.Count)) as Excel.Range;
            //bodyRange.AutoFilter(1, Type.Missing, Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);


            // startRow = rowNumber + 1; startColumn = columnNumber;
            //string lastColumn = clsManageSheet.getColumnName(dt.Columns.Count - 1);
            //Excel.Range bodyRange = worksheet.get_Range("B" + Convert.ToString(rowNumber), lastColumn + Convert.ToString(11 + dtPromoRows.Rows.Count)) as Excel.Range;
            //bodyRange.AutoFilter(1, Type.Missing, Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);


            //for (int i = 0; i < dt.Columns.Count - 1; i++)
            //{
            //    worksheet.Cells[startRow, startColumn].Value = dt.Columns[i].ColumnName;
            //    startColumn++;
            //}

            //object[,] arr = new object[dt.Rows.Count, dt.Columns.Count];
            //for (int r = 0; r < dt.Rows.Count; r++)
            //{
            //    DataRow dr = dt.Rows[r];
            //    for (int c = 0; c < dt.Columns.Count; c++)
            //    {
            //        arr[r, c] = dr[c];
            //    }
            //}

            //Excel.Range c1 = (Excel.Range)worksheet.Cells[12, 2];
            //Excel.Range c2 = (Excel.Range)worksheet.Cells[12 + dt.Rows.Count - 1, dt.Columns.Count];
            //Excel.Range range = worksheet.get_Range(c1, c2);

            //range.Value = arr;


            //string lastColumn = clsManageSheet.getColumnName(dt.Columns.Count - 1);
            //Excel.Range bodyRange = worksheet.get_Range("B" + Convert.ToString(startRow), lastColumn + Convert.ToString(11 + dtPromoRows.Rows.Count)) as Excel.Range;
            //bodyRange.AutoFilter(1, Type.Missing, Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);
            #endregion
        }
        #endregion

        #region Applying Back Ground Color To Sheet
        private void applyingBackGroundColorToSheet(Worksheet sheet, int startRow, long endRow, string startColumn, string endColumn)
        {
            sheet.Select();

            Excel.Range headerRange = sheet.get_Range(startColumn + Convert.ToString(startRow), endColumn + Convert.ToString(startRow)) as Excel.Range;
            headerRange.Interior.ColorIndex = 1;
            headerRange.Font.Bold = true;
            headerRange.RowHeight = 20;
            headerRange.Font.ColorIndex = 2;

            Excel.Range range = sheet.get_Range(startColumn + Convert.ToString(startRow + 1), endColumn + Convert.ToString(endRow)) as Excel.Range;
            range.Interior.ColorIndex = 2;



            if (sheet.Name == clsInformation.PROMO_INPUT_TOOL)
            {

                Excel.Borders border = range.Borders;
                border[Excel.XlBordersIndex.xlInsideHorizontal].Color = ColorTranslator.FromHtml("#C0C0C0");
                border[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;


                Excel.Range rangeEdit = sheet.get_Range("D" + Convert.ToString(startRow + 1), "O" + Convert.ToString(endRow)) as Excel.Range;
                rangeEdit.Interior.ColorIndex = 19;
                Excel.Range rangeEdit1 = sheet.get_Range("P" + Convert.ToString(startRow + 1), endColumn + Convert.ToString(endRow)) as Excel.Range;
                rangeEdit1.Interior.Color = ColorTranslator.FromHtml("#DDEBF7");//Light Blue

                endColumn = "C";
            }

            int i;
            for (i = startRow + 1; i <= endRow; i++)
            {
                Excel.Range rangeBackgroundColor = sheet.get_Range(startColumn + Convert.ToString(i), endColumn + Convert.ToString(i)) as Excel.Range;
                rangeBackgroundColor.Interior.Color = ColorTranslator.FromHtml("#fdd49b");

                i = i + 1;
            }
        }
        #endregion

        #region Applying Formulas To Promo Input Tool Sheet
        private void applyingFormulasToPromoInputToolSheet(Worksheet worksheet, DataTable dt, int startRow, int startColumn)
        {
            int rowNumber = startRow, columnNumber = startColumn;
            // Performance improved 
            for (int i = 0; i < dt.Columns.Count - 1; i++)
            {

                string value = Convert.ToString(dt.Columns[i].ColumnName);
                string formula = (from DataRow dr in dtPromoFormulas.Rows
                                  where Convert.ToString(dr["Column"]).ToLower() == value.ToLower()
                                  select Convert.ToString(dr["Formula"])).FirstOrDefault();

                startRow = rowNumber + 1;
                if (formula != "" && formula != null)
                {
                    try
                    {
                        Excel.Range rngFormula = worksheet.get_Range(clsManageSheet.getColumnName(startColumn) + startRow, clsManageSheet.getColumnName(startColumn) + (startRow + dtPromoRows.Rows.Count - 1)) as Excel.Range;
                        rngFormula.Formula = formula;
                    }
                    catch (Exception ex)
                    {
                        FAST.errorLog(ex.Message + ex.StackTrace, "Promotion_Planning_applyingFormulasToPromoInputToolSheet");
                    }
                }

                startColumn++;
            }
            #region Commented code
            // Commented by 
            /*
                        while (Convert.ToString(worksheet.Cells[startRow, startColumn].Value) != "" && Convert.ToString(worksheet.Cells[startRow, startColumn].Value) != null)
                        {
                            string value = Convert.ToString(worksheet.Cells[startRow, startColumn].Value);
                            string formula = (from DataRow dr in dtPromoFormulas.Rows
                                              where (string)dr["Column"] == value
                                              select Convert.ToString(dr["Formula"])).FirstOrDefault();

                            startRow = rowNumber + 1;

                            if (formula != "" && formula != null)
                            {
                                try
                                {
                                    Excel.Range rngFormula = worksheet.get_Range(clsManageSheet.getColumnName(startColumn) + startRow, clsManageSheet.getColumnName(startColumn) + (startRow + dtPromoRows.Rows.Count - 1)) as Excel.Range;
                                    rngFormula.Formula = formula;
                                }
                                catch (Exception ex)
                                {
                                    FAST.errorLog(ex.Message + ex.StackTrace, "Promotion_Planning_applyingFormulasToPromoInputToolSheet");
                                }

                            }

                            startRow = rowNumber; startColumn++;

                        }
                        */
            #endregion
        }
        #endregion

        #region Reset DataTable Columns Order
        public DataTable resetDataTableColumnsOrder(DataTable table, String columnNames)
        {
            string[] headers = columnNames.Replace(" ", string.Empty).Split(',');

            for (int i = 0; i < headers.Length; i++)
            {
                table.Columns[headers[i]].SetOrdinal(i);
            }

            return table;
        }
        #endregion

        #region Data Validations For PromoInputTemplate
        private void dataValidationsForPromoInputTemplate(Worksheet worksheet, int startRow, int endRow)
        {
            int colIndex = 4;


            worksheet.Select();
            worksheet.Unprotect();

            #region Comments Over Header
            commentsForSheet(worksheet, 11, "F", "Select 'Blended - VDP Mix' for TCPUCodename unless the promotion is only for a specific config.");
            commentsForSheet(worksheet, 11, "L", "Please use local currency");
            commentsForSheet(worksheet, 11, "M", "Lift	Among Lift, Elasticity and IncrementalUnitsOverride, you should input only one of the three. The precedence rule is: Lift -> Elasticity->IncrementalUnitsOverride.Baselines units is required if you use the lift input.");
            commentsForSheet(worksheet, 11, "N", "Among Lift, Elasticity and IncrementalUnitsOverride, you should input only one of the three. The precedence rule is: Lift -> Elasticity ->  IncrementalUnitsOverride.  Baselines units is required if you use the elasticity input.");
            commentsForSheet(worksheet, 11, "O", "Total Incremental units during the promo period (not daily). For non-TPR promotions with NO baseline units, it is mandatory to input IncrementalUnitsOverride. Among Lift, Elasticity and IncrementalUnitsOverride, you should input only one of the three.");
            commentsForSheet(worksheet, 11, "Q", "Optional input: use if default AmazonFundingSplit is incorrect. For online, default is 100%; for offline, default % varies by product lines; for broadcast promo type, it is default to be 0%. ");
            commentsForSheet(worksheet, 11, "R", "For TPR promotions, default BaselineUnits is the program total baseline units during the promo period. For non-TPR promotions (including Targeted TPR), there's no baseline units in default.");
            commentsForSheet(worksheet, 11, "S", "Please input BaselineUnitsOverride if you use a lift or elasticity input but there's no baselineUnits or if baselineUnits is inaccurate.");
            commentsForSheet(worksheet, 11, "U", "Optional input: use only when default TCPU mapping program is incorrect");
            commentsForSheet(worksheet, 11, "W", "Default MSRP before VAT (Local currency)");
            commentsForSheet(worksheet, 11, "V", "Default MSRP after VAT (USD)");
            commentsForSheet(worksheet, 11, "X", "Optional input: Please use local currency. Value entered will replace INTLMSRP");
            commentsForSheet(worksheet, 11, "Z", "Optional input");
            commentsForSheet(worksheet, 11, "AB", "Optional input:offline/wholesale discount in USD currency. Usually a negative number. Value entered will replace the default sales discount.");
            commentsForSheet(worksheet, 11, "AD", "Optional input: RPU excluding lifetime promo risk. In USD currency. Value will replace default baseline RPU. ");
            commentsForSheet(worksheet, 11, "AF", "Optional input: PPU excluding lifetime promo risk. In USD currency. Value will replace default baseline PPU.");
            commentsForSheet(worksheet, 11, "AH", "Optional input: estimated CMA DSI for device sold at FULL PRICE. In USD currency. Value will replace default baseline CMA DSI.");
            commentsForSheet(worksheet, 11, "AJ", "Optional input: Non CMA DSI such as prime propensity, accessories and hawfire. In USD currency.");
            commentsForSheet(worksheet, 11, "AL", "Optional input");
            commentsForSheet(worksheet, 11, "AM", "Optional input: total additional cost for the promotion (other than promo discount)  in USD currency.");
            #endregion


            while (Convert.ToString(worksheet.Cells[startRow, colIndex].Value) != null && Convert.ToString(worksheet.Cells[startRow, colIndex].Value) != "")
            {
                string ColumnHeader = Convert.ToString(worksheet.Cells[startRow, colIndex].Value);

                if (ColumnHeader == "Discount" || ColumnHeader == "Lift" || ColumnHeader == "IncrementalUnitsOverride" || ColumnHeader == "AmazonFundingSplitOverride" || ColumnHeader == "Elasticity")
                {
                    Excel.Range validationRange = null;
                    switch (ColumnHeader)
                    {
                        case "Discount":
                            validationRange = worksheet.get_Range("L" + Convert.ToString(startRow + 1), "L" + Convert.ToString(endRow)) as Excel.Range;
                            validationRange.Validation.Add(Excel.XlDVType.xlValidateDecimal, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlGreaterEqual, 0);
                            validationRange.Validation.ErrorMessage = "Please Give inputs Greater than or Equal to 0";
                            break;
                        case "Lift":
                            validationRange = worksheet.get_Range("M" + Convert.ToString(startRow + 1), "M" + Convert.ToString(endRow)) as Excel.Range;
                            validationRange.Validation.Add(Excel.XlDVType.xlValidateDecimal, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlGreaterEqual, 1);//"999999999999999"
                            validationRange.Validation.ErrorMessage = "Please Give inputs Greater than or Equal to 1";
                            break;
                        case "IncrementalUnitsOverride":
                            validationRange = worksheet.get_Range("O" + Convert.ToString(startRow + 1), "O" + Convert.ToString(endRow)) as Excel.Range;
                            validationRange.Validation.Add(Excel.XlDVType.xlValidateDecimal, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlGreaterEqual, 0);
                            validationRange.Validation.ErrorMessage = "Please Give inputs Greater than or Equal to 0";
                            break;
                        case "AmazonFundingSplitOverride":
                            validationRange = worksheet.get_Range("Q" + Convert.ToString(startRow + 1), "Q" + Convert.ToString(endRow)) as Excel.Range;
                            //validationRange.NumberFormat = "##0.00\\%";
                            validationRange.Validation.Add(Excel.XlDVType.xlValidateDecimal, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, 0, 1);//"999999999999999"
                            validationRange.Validation.ErrorMessage = "Please Give inputs between 0 and 1";
                            break;
                        case "Elasticity":
                            validationRange = worksheet.get_Range("N" + Convert.ToString(startRow + 1), "N" + Convert.ToString(endRow)) as Excel.Range;
                            validationRange.Validation.Add(Excel.XlDVType.xlValidateDecimal, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlGreater, 0);//"999999999999999"
                            validationRange.Validation.ErrorMessage = "Please Give inputs Greater than or Equal to 1";
                            break;
                    }

                }

                colIndex++;
            }

        }
        #endregion

        #region Comments For Sheet
        private void commentsForSheet(Worksheet worksheet, int startRow, string endRow, string comment)
        {
            Excel.Comment rng = worksheet.Cells[startRow, endRow].AddComment(comment);
            rng.Shape.TextFrame.AutoSize = true;


        }
        #endregion

        #region Making Columns Read Only
        private void makingColumnsReadOnly(Worksheet worksheet, DataTable dt, DataTable dtPromoFormulas, int bodyStartRow, long lastRow)
        {

            var list = dtPromoFormulas.Rows.OfType<DataRow>()
                             .Select(dr => dr.Field<string>("Column")).ToList();
            list = list.ConvertAll(a => a.ToUpper());
            int colIndex = 4;


            //     while (Convert.ToString(worksheet.Cells[bodyStartRow, colIndex].Value) != null && Convert.ToString(worksheet.Cells[11, colIndex].Value) != "")
            // {

            for (int i = 0; i < dt.Columns.Count - 1; i++)
            {
                string columnName = Convert.ToString(worksheet.Cells[11, colIndex].Value);

                columnName = columnName != null ? columnName.ToUpper() : "";

                var getCount = list.Where(a => a == columnName).ToList();

                string getColumn = clsManageSheet.getColumnName(colIndex);
                Excel.Range readonlyRange = worksheet.get_Range(getColumn + Convert.ToString(bodyStartRow + 1), getColumn + Convert.ToString(lastRow)) as Excel.Range;
                if (getCount.Count != 0) // which means row is available in the table and to be made readonly
                {
                    readonlyRange.Interior.Color = ColorTranslator.FromHtml("#8c8c8c");
                }
                else
                {
                    readonlyRange.Locked = false;
                    // readonlyRange.FormulaHidden = false;
                }

                colIndex++;
            }

            // starts


        }
        #endregion

        #region Resetting Column Width
        private void resettingColumnWidth(Worksheet ExcelSheet, int startColumnNumber, int startRowNumber)
        {
            while (Convert.ToString(ExcelSheet.Cells[startRowNumber, startColumnNumber].Value) != null && Convert.ToString(ExcelSheet.Cells[startRowNumber, startColumnNumber].Value) != "")
            {
                string value = Convert.ToString(ExcelSheet.Cells[startRowNumber, startColumnNumber].Value);
                int length = value.Length;

                switch (value)
                {
                    case "TCPUCodename":
                    case "PROGRAM_TCPU_MAPPING":
                    case "TOTAL_PROMO_UNITS_FORECAST":
                    case "COUNTRY":
                        ExcelSheet.Cells[startRowNumber, startColumnNumber].ColumnWidth = length + 12;
                        break;
                    case "DeviceType":
                    case "Country":
                    case "CURRENCY":
                    case "MSRP":
                    case "country":
                    case "program":
                    case "TCPU_Codename":
                    case "date":
                    case "DEVICE_TYPE":
                    case "RUN_DATE":
                        ExcelSheet.Cells[startRowNumber, startColumnNumber].ColumnWidth = length + 17;
                        break;
                    case "VLOOKUP":
                    case "DESCRIPTION":
                        ExcelSheet.Cells[startRowNumber, startColumnNumber].ColumnWidth = length + 46;
                        break;
                    case "PROMOTION_ID":
                        ExcelSheet.Cells[startRowNumber, startColumnNumber].ColumnWidth = length + 30;
                        break;
                    default:
                        ExcelSheet.Cells[startRowNumber, startColumnNumber].ColumnWidth = length + 7;
                        break;

                }
                startColumnNumber++;
            }
        }
        #endregion

        #region Add Dropdown List To Columns
        private static void addDropdownListToColumns(Worksheet worksheet, int startColumnNumber, int startRowNumber, long endRow,


                                     string deviceType, string programList, string tcpuCodename, string Channel, string PromoType)
        {

            int ProgramList = 0, tcpuCodeNameList = 0, channelList = 0, promotype = 0;


            for (int i = 0; i < dsDeviceTypes.Tables.Count; i++)
            {
                if (dsDeviceTypes.Tables[i].TableName == "ProgramList" || dsDeviceTypes.Tables[i].TableName == "Program")
                    ProgramList = i;
                if (dsDeviceTypes.Tables[i].TableName == "TcpuCodeNameList" || dsDeviceTypes.Tables[i].TableName == "TcpuCodeName")
                {
                    tcpuCodeNameList = i;
                    tcpuCodeNameListTableId = tcpuCodeNameList;
                }

                if (dsDeviceTypes.Tables[i].TableName == "Channel")
                    channelList = i;
                if (dsDeviceTypes.Tables[i].TableName == "Promo")
                    promotype = i;


            }

            // First setting the dropdowns for both the PromoType and Channel as they will be only one

            if (Convert.ToString(dsDeviceTypes.Tables[promotype].Rows[0].ItemArray[0]) != null && Convert.ToString(dsDeviceTypes.Tables[promotype].Rows[0].ItemArray[0]) != "")
            {
                Excel.Range promoTypeRange = worksheet.get_Range(PromoType + Convert.ToString(startRowNumber), PromoType + Convert.ToString(endRow)) as Excel.Range;
                promoTypeRange.Validation.Delete();
                promoTypeRange.Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, Convert.ToString(dsDeviceTypes.Tables[promotype].Rows[0].ItemArray[0]));
            }


            // Channel Type
            if (Convert.ToString(dsDeviceTypes.Tables[channelList].Rows[0].ItemArray[0]) != null && Convert.ToString(dsDeviceTypes.Tables[channelList].Rows[0].ItemArray[0]) != "")
            {
                Excel.Range channelRange = worksheet.get_Range(Channel + Convert.ToString(startRowNumber), Channel + Convert.ToString(endRow)) as Excel.Range;
                channelRange.Validation.Delete();
                channelRange.Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, Convert.ToString(dsDeviceTypes.Tables[channelList].Rows[0].ItemArray[0]));
            }



            for (int i = 0; i < dsDeviceTypes.Tables[ProgramList].Rows.Count; i++)
            {

                for (int j = 0; j < dsDeviceTypes.Tables[ProgramList].Columns.Count; j++)
                {
                    startRowNumber = 12;

                    string[] programListColumnName = Convert.ToString(dsDeviceTypes.Tables[ProgramList].Rows[i].ItemArray[j]).Split('|');

                    while (startRowNumber <= endRow)
                    {
                        string deviceTypeFromSheet = Convert.ToString(worksheet.Cells[startRowNumber, deviceType].Value);

                        if (deviceTypeFromSheet == programListColumnName[0])
                        {
                            Excel.Range programListRange = worksheet.get_Range(programList + Convert.ToString(startRowNumber), programList + Convert.ToString(startRowNumber)) as Excel.Range;
                            programListRange.Validation.Delete();
                            programListRange.Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, programListColumnName[1]);
                        }

                        startRowNumber++;
                    }
                }

            }

            //startRowNumber = 12;

            for (int i = 0; i < dsDeviceTypes.Tables[tcpuCodeNameList].Rows.Count; i++)
            {

                for (int j = 0; j < dsDeviceTypes.Tables[tcpuCodeNameList].Columns.Count; j++)
                {
                    startRowNumber = 12;

                    string[] tcpuCodeNameListColumnName = Convert.ToString(dsDeviceTypes.Tables[tcpuCodeNameList].Rows[i].ItemArray[j]).Split('|');


                    while (startRowNumber <= endRow)
                    {
                        string deviceTypeFromSheet = Convert.ToString(worksheet.Cells[startRowNumber, deviceType].Value);

                        if (deviceTypeFromSheet == tcpuCodeNameListColumnName[0])
                        {
                            Excel.Range tcpuCodenameRange = worksheet.get_Range(tcpuCodename + Convert.ToString(startRowNumber), tcpuCodename + Convert.ToString(startRowNumber)) as Excel.Range;
                            tcpuCodenameRange.Validation.Delete();
                            tcpuCodenameRange.Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, tcpuCodeNameListColumnName[1]);
                        }

                        startRowNumber++;
                    }


                }


            }

        }

        #endregion

        // Refactored Code base ends here by Nihar

        #region promotion Input Tool Upload
        public static bool verifyDownloadForPromoUpload()
        {
            ExcelTool.Worksheet verifyPromo = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[clsInformation.PROMO_INPUT_TOOL]);
            if (verifyPromo.UsedRange.Rows.Count < 13 /*|| FAST._tcpu.UsedRange.Rows.Count < 13 || FAST._vdp.UsedRange.Rows.Count < 13*/)
            {
                FAST.displayAlerts(clsInformation.verifyDownloadforUpload, 1);
                return false;
            }
            return true;

        }
        
        public static string promotionPlanningUpload(string userName, string selectedProcess, string countryId, string deviceTypeId)
        {
            Worksheet worksheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[clsInformation.PROMO_INPUT_TOOL]);

            dtPromoMaster = dsPromotions.Tables[promoMasterTableName];
            dtPromoRows = dsPromotions.Tables[promoRowsTableName];

            string[] columns = dtPromoMaster.Rows[0]["PromoHeader"].ToString().ReplaceNewLine().Split(',');
            string _promoLastColumnName = clsManageSheet.getColumnName(columns.Length + 1);

            if (!verifyValidationsforPromoUpload("D", "J", "K", "L", "M", "N", "O", "Q", 12, (11 + dtPromoRows.Rows.Count), worksheet, _promoLastColumnName))
            {
                return "";//"Upload Failed due to validation failures.";
            }

            _promoLastColumnName += Convert.ToString(11);

            Excel.Range header = worksheet.get_Range("D11", _promoLastColumnName) as Excel.Range;

            int rowCount = 11 + dtPromoRows.Rows.Count;

            DataTable dt = makeDataTableFromPromoRange(header, worksheet, rowCount);

            var sbPromo = new StringBuilder();

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string val = null;
                int j = 0;

                // Need to verify for Time Stamp
                for (j = 0; j < dt.Columns.Count - 1; j++)
                {
                    if (dt.Columns[j].ColumnName == "StartDate" || dt.Columns[j].ColumnName == "EndDate")
                    {
                        string cellValue = Convert.ToString(dt.Rows[i].ItemArray[j]) == "" ? "null" : Convert.ToString(dt.Rows[i].ItemArray[j]);

                        if (cellValue != "null")
                        {
                            DateTime oDate = DateTime.Parse(cellValue);
                            val += "\"" + oDate.ToString("MM/dd/yyyy") + "\",";

                        }

                        else
                            val += "\"" + cellValue + "\",";



                    }
                    else
                    {
                        string cellValue = Convert.ToString(dt.Rows[i].ItemArray[j]) == "" ? "null" : Convert.ToString(dt.Rows[i].ItemArray[j]);
                        val += "\"" + cellValue + "\",";
                        //val.Replace("\"null\"", "null");

                    }


                }

                val = val.TrimEnd(',');


                string pcIdValue = Convert.ToString(dt.Rows[i].ItemArray[j]) == "" ? "null" : Convert.ToString(dt.Rows[i].ItemArray[j]);
                val = "\"" + pcIdValue + "\"," + val;


                sbPromo.AppendFormat("({0}),", val);



            }

            var sbPromoDatatoUpload = sbPromo.ToString().TrimEnd(',');



            //For Uploading.
            var response = FASTWebServiceAdapter.sendUploadDataForPromotions(userName, selectedProcess, countryId, deviceTypeId, FAST._txtProcess, sbPromoDatatoUpload);
            return response.Tables[0].Rows[0]["IsSuccess"].ToString();
        }

        private static DataTable makeDataTableFromPromoRange(Excel.Range headerRange, ExcelTool.Worksheet sheet, long usedRange)
        {
            DataTable table = new DataTable();
            foreach (Excel.Range col in headerRange)
            {

                table.Columns.Add(col.Text);

            }
            for (long l = 12; l <= usedRange; l++)
            {

                DataRow row = table.NewRow();
                int headerCount = headerRange.Count + 4;

                for (int i = 4; i < headerCount; i++)
                {
                    row[i - 4] = sheet.Cells[l, i].Value;
                }

                table.Rows.Add(row);
            }

            return table;
        }



        #endregion 

        #region Upload Validations For Promo Input Template
        public static bool verifyValidationsforPromoUpload(string VersionColumn, string startDate,
                                                           string endDate, string discount,
                                                           string lift, string Elasticity, string incrementalUnits, string amazonFundingSplit,
                                                           int startRow, long EndRow,
                                                            ExcelTool.Worksheet sheet, string endColumn)
        {
            sheet.Select();

            int promoTypeColNumber = 10, channelColNumber = 9, programColNumber = 5, tcpuCodeNameColNumber = 8;

            unProtect(sheet);

            Excel.Range changeFontColorForRange = sheet.get_Range("B" + Convert.ToString(startRow), endColumn + Convert.ToString(EndRow)) as Excel.Range;
            changeFontColorForRange.Font.ColorIndex = 1;

            Excel.Range versionRange = sheet.get_Range(VersionColumn + Convert.ToString(startRow), VersionColumn + Convert.ToString(EndRow)) as Excel.Range;
            versionRange.Interior.ColorIndex = 19;

            Excel.Range startDateRange = sheet.get_Range(startDate + Convert.ToString(startRow), startDate + Convert.ToString(EndRow)) as Excel.Range;
            startDateRange.Interior.ColorIndex = 19;

            Excel.Range endDateRange = sheet.get_Range(endDate + Convert.ToString(startRow), endDate + Convert.ToString(EndRow)) as Excel.Range;
            endDateRange.Interior.ColorIndex = 19;

            Excel.Range discountRange = sheet.get_Range(discount + Convert.ToString(startRow), discount + Convert.ToString(EndRow)) as Excel.Range;
            discountRange.Interior.ColorIndex = 19;

            Excel.Range liftRange = sheet.get_Range(lift + Convert.ToString(startRow), lift + Convert.ToString(EndRow)) as Excel.Range;
            liftRange.Interior.ColorIndex = 19;

            Excel.Range elasticityRange = sheet.get_Range(Elasticity + Convert.ToString(startRow), Elasticity + Convert.ToString(EndRow)) as Excel.Range;
            elasticityRange.Interior.ColorIndex = 19;

            Excel.Range incrementalUnitsRange = sheet.get_Range(incrementalUnits + Convert.ToString(startRow), incrementalUnits + Convert.ToString(EndRow)) as Excel.Range;
            incrementalUnitsRange.Interior.ColorIndex = 19;

            Excel.Range amazonFundingSplitRange = sheet.get_Range(amazonFundingSplit + Convert.ToString(startRow), amazonFundingSplit + Convert.ToString(EndRow)) as Excel.Range;

            Excel.Range programRange = sheet.get_Range(clsManageSheet.getColumnName(programColNumber) + Convert.ToString(startRow), clsManageSheet.getColumnName(programColNumber) + Convert.ToString(EndRow)) as Excel.Range;
            programRange.Interior.ColorIndex = 19;

            Excel.Range tcpuCodeRange = sheet.get_Range(clsManageSheet.getColumnName(tcpuCodeNameColNumber) + Convert.ToString(startRow), clsManageSheet.getColumnName(tcpuCodeNameColNumber) + Convert.ToString(EndRow)) as Excel.Range;
            tcpuCodeRange.Interior.ColorIndex = 19;

            Excel.Range channelRange = sheet.get_Range(clsManageSheet.getColumnName(channelColNumber) + Convert.ToString(startRow), clsManageSheet.getColumnName(channelColNumber) + Convert.ToString(EndRow)) as Excel.Range;
            channelRange.Interior.ColorIndex = 19;

            Excel.Range promotypeRange = sheet.get_Range(clsManageSheet.getColumnName(promoTypeColNumber) + Convert.ToString(startRow), clsManageSheet.getColumnName(promoTypeColNumber) + Convert.ToString(EndRow)) as Excel.Range;
            promotypeRange.Interior.ColorIndex = 19;

            protectSheet(sheet);

            while (startRow < EndRow)
            {
                Excel.Range highlightRange = sheet.get_Range("B" + Convert.ToString(startRow), endColumn + Convert.ToString(startRow)) as Excel.Range;

                string promoTypeColVal = Convert.ToString(sheet.Cells[startRow, promoTypeColNumber].Value);
                string channelColVal = Convert.ToString(sheet.Cells[startRow, channelColNumber].Value);
                string programColVal = Convert.ToString(sheet.Cells[startRow, programColNumber].Value);
                string tcpuCodeNameColVal = Convert.ToString(sheet.Cells[startRow, tcpuCodeNameColNumber].Value);


                string value = Convert.ToString(sheet.Cells[startRow, VersionColumn].Value);

                if (value != "" && value != null)
                {
                    #region  Validate dropdown columns

                    Excel.Range programColRange = sheet.get_Range("E" + Convert.ToString(startRow), "E" + Convert.ToString(startRow)) as Excel.Range;

                    if (string.IsNullOrWhiteSpace(programColVal) || (!string.IsNullOrWhiteSpace(programColVal) && !programColRange.Validation.Value))// 
                    {
                        unProtect(sheet);

                        sheet.Cells[startRow, programColNumber].Interior.Color = ColorTranslator.FromHtml("#ffd700");
                        sheet.Cells[startRow, programColNumber].Select();
                        highlightRange.Font.Color = ColorTranslator.FromHtml("#FF0000");

                        protectSheet(sheet);

                        if (string.IsNullOrWhiteSpace(programColVal))
                            FAST.displayAlerts("Program Column should not contain Blank cells for Uploading.", 2);
                        else
                            FAST.displayAlerts("Program Column has invalid data. Please select data from dropdown list only.", 2);

                        return false;

                    }

                    Excel.Range tcpuCodeNameColRange = sheet.get_Range("H" + Convert.ToString(startRow), "H" + Convert.ToString(startRow)) as Excel.Range;
                    if (!tcpuCodeNameColRange.Validation.Value)//string.IsNullOrWhiteSpace(tcpuCodeNameColVal) || (!string.IsNullOrWhiteSpace(tcpuCodeNameColVal) &&
                    {
                        unProtect(sheet);

                        sheet.Cells[startRow, tcpuCodeNameColNumber].Interior.Color = ColorTranslator.FromHtml("#ffd700");
                        sheet.Cells[startRow, tcpuCodeNameColNumber].Select();
                        highlightRange.Font.Color = ColorTranslator.FromHtml("#FF0000");

                        protectSheet(sheet);

                        if (string.IsNullOrWhiteSpace(tcpuCodeNameColVal))
                            FAST.displayAlerts("Tcpu Code Name Column should not contain Blank cells for Uploading.", 2);
                        else
                            FAST.displayAlerts("Tcpu Code Name Column has invalid data. Please select data from dropdown list only.", 2);

                        return false;

                    }
                    Excel.Range channelColRange = sheet.get_Range("I" + Convert.ToString(startRow), "I" + Convert.ToString(startRow)) as Excel.Range;
                    if (!channelColRange.Validation.Value)//string.IsNullOrWhiteSpace(channelColVal) || (!string.IsNullOrWhiteSpace(channelColVal) && 
                    {
                        unProtect(sheet);

                        sheet.Cells[startRow, channelColNumber].Interior.Color = ColorTranslator.FromHtml("#ffd700");
                        sheet.Cells[startRow, channelColNumber].Select();
                        highlightRange.Font.Color = ColorTranslator.FromHtml("#FF0000");

                        protectSheet(sheet);
                        if (!channelColRange.Validation.Value)
                            FAST.displayAlerts("Channel Column has invalid data. Please select data from dropdown list only.", 2);

                        return false;

                    }

                    Excel.Range promotypeColRange = sheet.get_Range("J" + Convert.ToString(startRow), "J" + Convert.ToString(startRow)) as Excel.Range;

                    if (!promotypeColRange.Validation.Value)//string.IsNullOrWhiteSpace(promoTypeColVal) || (!string.IsNullOrWhiteSpace(promoTypeColVal) && 
                    {
                        unProtect(sheet);

                        sheet.Cells[startRow, promoTypeColNumber].Interior.Color = ColorTranslator.FromHtml("#ffd700");
                        sheet.Cells[startRow, promoTypeColNumber].Select();
                        highlightRange.Font.Color = ColorTranslator.FromHtml("#FF0000");

                        protectSheet(sheet);
                        if (!promotypeColRange.Validation.Value)//string.IsNullOrWhiteSpace(promoTypeColVal) || (!string.IsNullOrWhiteSpace(promoTypeColVal) && 
                            FAST.displayAlerts("PromoType Column  has invalid data. Please select data from dropdown list only.", 2);

                        return false;

                    }



                    #endregion

                    #region Start Date
                    // For Start Date
                    value = Convert.ToString(sheet.Cells[startRow, startDate].Value);

                    if (value == "" || value == null)
                    {
                        unProtect(sheet);

                        sheet.Cells[startRow, startDate].Interior.Color = ColorTranslator.FromHtml("#ffd700");
                        sheet.Cells[startRow, startDate].Select();
                        highlightRange.Font.Color = ColorTranslator.FromHtml("#FF0000");

                        protectSheet(sheet);

                        FAST.displayAlerts("Start Date Column should not contain Blank cells for Uploading.", 2);

                        return false;
                    }
                    try
                    {
                        string iDate = value;
                        DateTime oDate = DateTime.Parse(iDate);
                    }
                    catch
                    {
                        unProtect(sheet);

                        sheet.Cells[startRow, startDate].Interior.Color = ColorTranslator.FromHtml("#ffd700");
                        sheet.Cells[startRow, startDate].Select();
                        highlightRange.Font.Color = ColorTranslator.FromHtml("#FF0000");

                        protectSheet(sheet);

                        FAST.displayAlerts("Invalid Date Format in Start Date Column - Please Verify.", 2);

                        return false;
                    }

                    #endregion

                    #region EndDate
                    // For End Date
                    value = Convert.ToString(sheet.Cells[startRow, endDate].Value);

                    if (value == "" || value == null)
                    {
                        unProtect(sheet);

                        sheet.Cells[startRow, endDate].Interior.Color = ColorTranslator.FromHtml("#ffd700");
                        sheet.Cells[startRow, endDate].Select();
                        highlightRange.Font.Color = ColorTranslator.FromHtml("#FF0000");

                        protectSheet(sheet);

                        FAST.displayAlerts("End Date should not contain Blank cells for Uploading.", 2);

                        return false;
                    }

                    try
                    {
                        string iDate = value;
                        DateTime oDate = DateTime.Parse(iDate);
                    }
                    catch
                    {
                        unProtect(sheet);

                        sheet.Cells[startRow, endDate].Interior.Color = ColorTranslator.FromHtml("#ffd700");
                        sheet.Cells[startRow, endDate].Select();
                        highlightRange.Font.Color = ColorTranslator.FromHtml("#FF0000");

                        protectSheet(sheet);

                        FAST.displayAlerts("Invalid Date Format in End Date Column - Please Verify.", 2);

                        return false;
                    }

                    #endregion

                    #region Discount Column

                    value = Convert.ToString(sheet.Cells[startRow, discount].Value);


                    if (value != "" || value != null)
                    {
                        if ((Convert.ToDouble(value) < 0))
                        {
                            unProtect(sheet);

                            sheet.Cells[startRow, discount].Interior.Color = ColorTranslator.FromHtml("#ffd700");
                            sheet.Cells[startRow, discount].Select();
                            highlightRange.Font.Color = ColorTranslator.FromHtml("#FF0000");

                            protectSheet(sheet);

                            FAST.displayAlerts("Values should not be less than 0 for Discount Column for Uploading..", 2);

                            return false;
                        }
                    }



                    #endregion

                    #region Lift

                    value = Convert.ToString(sheet.Cells[startRow, lift].Value);

                    if (value != null)
                    {
                        if (Convert.ToDouble(value) < 1)
                        {
                            unProtect(sheet);

                            sheet.Cells[startRow, lift].Interior.Color = ColorTranslator.FromHtml("#ffd700");
                            sheet.Cells[startRow, lift].Select();
                            highlightRange.Font.Color = ColorTranslator.FromHtml("#FF0000");

                            protectSheet(sheet);

                            FAST.displayAlerts("Values should not be less than 1 for Lift Column for Uploading.", 2);

                            return false;
                        }

                    }

                    #endregion

                    #region Elasticity

                    value = Convert.ToString(sheet.Cells[startRow, Elasticity].Value);

                    if (value != null)
                    {
                        if (Convert.ToDouble(value) < 0)
                        {
                            unProtect(sheet);

                            sheet.Cells[startRow, Elasticity].Interior.Color = ColorTranslator.FromHtml("#ffd700");
                            sheet.Cells[startRow, Elasticity].Select();
                            highlightRange.Font.Color = ColorTranslator.FromHtml("#FF0000");

                            protectSheet(sheet);

                            FAST.displayAlerts("Values should not be less than 0 for Elasticity Column for Uploading.", 2);

                            return false;
                        }
                    }


                    #endregion

                    #region Incremental Units

                    value = Convert.ToString(sheet.Cells[startRow, incrementalUnits].Value);

                    if (value != null)
                    {
                        if (Convert.ToDouble(value) < 0)
                        {
                            unProtect(sheet);

                            sheet.Cells[startRow, incrementalUnits].Interior.Color = ColorTranslator.FromHtml("#ffd700");
                            sheet.Cells[startRow, incrementalUnits].Select();
                            highlightRange.Font.Color = ColorTranslator.FromHtml("#FF0000");

                            protectSheet(sheet);

                            FAST.displayAlerts("Values should not be less than 0 for Incremental Units Overrride Column for Uploading.", 2);

                            return false;
                        }
                    }


                    #endregion

                    #region Amazon Funding Split Override

                    value = Convert.ToString(sheet.Cells[startRow, amazonFundingSplit].Value);

                    if (value != null)
                    {
                        //if (Convert.ToInt32(value) < 0 || Convert.ToInt32(value) > 1)
                        if (Convert.ToDouble(value) < 0 || Convert.ToDouble(value) > 1)
                        {
                            unProtect(sheet);

                            sheet.Cells[startRow, amazonFundingSplit].Interior.Color = ColorTranslator.FromHtml("#ffd700");
                            sheet.Cells[startRow, amazonFundingSplit].Select();
                            highlightRange.Font.Color = ColorTranslator.FromHtml("#FF0000");

                            protectSheet(sheet);

                            FAST.displayAlerts("Values should not be in between 0 and 1 for Amazon Funding Split Override Column for Uploading.", 2);

                            return false;
                        }
                    }


                    #endregion
                }


                #region Validate dropdown columns






                #endregion


                startRow++;
            }

            return true;
        }
        #endregion

        #region Protect Sheet

        /// <summary>
        /// This Method is used to Protect the Sheet
        /// </summary>
        /// <param name="worksheet"></param>

        public static void protectSheet(Microsoft.Office.Tools.Excel.Worksheet worksheet)
        {
            worksheet.Protect("PromoTool", worksheet.ProtectDrawingObjects,
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

        #region Unprotect Sheet

        /// <summary>
        /// This Method is used to unprotect the sheet
        /// </summary>
        /// <param name="visiblesheet"></param>

        public static void unProtect(Microsoft.Office.Tools.Excel.Worksheet visiblesheet)
        {
            visiblesheet.Unprotect("PromoTool");
        }
        #endregion

        #region Offline
        public static void uploadButtonClickforPromoOffline(bool isDownloadEnabled, bool isUploadEnabled, bool isTCPUVDPEnabled, bool isBransonEnabled)
        {
            Excel.Worksheet worksheet = null;
            clsManageSheet.buildSheet(ref worksheet, clsInformation.referencePromo);
            worksheet.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;

            //ClsPromotions.promotionsBeforeClose();
            //promotionsBeforeClose();

            int startRowIndex = 1, colValue = 1;
            if (worksheet == null && worksheet.Name != clsInformation.referenceDataSheet)
                throw new Exception("The Sheet is Not Validated For Data Building");
            //added by praveen1

            worksheet.Cells[startRowIndex, colValue++] = "Process";
            worksheet.Cells[startRowIndex, colValue++] = "ProcessId";
            worksheet.Cells[startRowIndex, colValue++] = "CountryValue";
            worksheet.Cells[startRowIndex, colValue++] = "DeviceTypeValue";
            worksheet.Cells[startRowIndex, colValue++] = "downloadCountryValue";
            worksheet.Cells[startRowIndex, colValue++] = "downloadDeviceTypeValue";
            worksheet.Cells[startRowIndex, colValue++] = "Saved File";
            worksheet.Cells[startRowIndex, colValue++] = "downloadCountryLabel";
            worksheet.Cells[startRowIndex, colValue++] = "downloadDeviceTypeLabel";
            worksheet.Cells[startRowIndex, colValue++] = "promoCountryRefreshTCPUVDPLabel";
            worksheet.Cells[startRowIndex, colValue++] = "promoDeviceRefreshTCPUVDPLabel";
            worksheet.Cells[startRowIndex, colValue++] = "promoCountryRefreshTCPUVDPValue";
            worksheet.Cells[startRowIndex, colValue++] = "promodeviceRefreshTCPUVDPValue";
            worksheet.Cells[startRowIndex, colValue++] = "promoCountryRefreshBransonLabel";
            worksheet.Cells[startRowIndex, colValue++] = "promoDeviceRefreshBransonLabel";
            worksheet.Cells[startRowIndex, colValue++] = "promoCountryRefreshBransonValue";
            worksheet.Cells[startRowIndex, colValue++] = "promodeviceRefreshBransonPValue";
            worksheet.Cells[startRowIndex, colValue++] = "deviceTypeFileSave"; // Added by Nihar on 11/30/2017
            worksheet.Cells[startRowIndex, colValue++] = "promoBodyEndRows";

            //added by Anwesh (22/08/2019)

            worksheet.Cells[startRowIndex, colValue++] = "isDownloadEnabled";
            worksheet.Cells[startRowIndex, colValue++] = "isUploadEnabled";
            worksheet.Cells[startRowIndex, colValue++] = "isRefreshVDPTCPUEnabled";
            worksheet.Cells[startRowIndex, colValue++] = "isRefreshBransonEnabled";
            worksheet.Cells[startRowIndex, colValue++] = clsInformation.AliasId;



            startRowIndex = 2; colValue = 1;

            worksheet.Cells[startRowIndex, colValue++] = FAST._txtProcess;
            worksheet.Cells[startRowIndex, colValue++] = FAST._valueProcess;
            worksheet.Cells[startRowIndex, colValue++] = FAST._promoCountryValue;
            worksheet.Cells[startRowIndex, colValue++] = FAST._promoDeviceTypeValue;
            worksheet.Cells[startRowIndex, colValue++] = FAST._promoDownloadCountryValueForOfflineOnline;
            worksheet.Cells[startRowIndex, colValue++] = FAST._promoDownloadDeviceTypeForOfflineOnline;
            worksheet.Cells[startRowIndex, colValue++] = "Country" + FAST._promoDownloadCountryValueForOfflineOnline + "_Device" + FAST._promoDownloadDeviceTypeForOfflineOnline + "_PromoPlanning.xml";
            worksheet.Cells[startRowIndex, colValue++] = FAST._promoCountryLabel;
            worksheet.Cells[startRowIndex, colValue++] = FAST._promoDeviceLabel;
            worksheet.Cells[startRowIndex, colValue++] = FAST._promoCountryRefreshTCPUVDPLable != null ? FAST._promoCountryRefreshTCPUVDPLable : FAST._promoCountryLabel;
            worksheet.Cells[startRowIndex, colValue++] = FAST._promodeviceRefreshTCPUVDPLabel != null ? FAST._promodeviceRefreshTCPUVDPLabel : FAST._promoDeviceLabel;
            worksheet.Cells[startRowIndex, colValue++] = FAST._promoCountryRefreshTCPUVDPValue != null ? FAST._promoCountryRefreshTCPUVDPValue : FAST._promoCountryValue;
            worksheet.Cells[startRowIndex, colValue++] = FAST._promodeviceRefreshTCPUVDPValue != null ? FAST._promodeviceRefreshTCPUVDPValue : FAST._promoDeviceTypeValue;
            worksheet.Cells[startRowIndex, colValue++] = FAST._promoCountryBransonRefreshLable != null ? FAST._promoCountryBransonRefreshLable : FAST._promoCountryLabel;
            worksheet.Cells[startRowIndex, colValue++] = FAST._promodeviceBransonRefreshLabel != null ? FAST._promodeviceBransonRefreshLabel : FAST._promoDeviceLabel; ;
            worksheet.Cells[startRowIndex, colValue++] = FAST._promoCountryBransonRefreshValue != null ? FAST._promoCountryBransonRefreshValue : FAST._promoCountryValue;
            worksheet.Cells[startRowIndex, colValue++] = FAST._promodeviceBransonRefreshValue != null ? FAST._promodeviceBransonRefreshValue : FAST._promoDeviceTypeValue;


            worksheet.Cells[startRowIndex, colValue++] = "Country" + FAST._promoDownloadCountryValueForOfflineOnline + "_Device" + FAST._promoDownloadDeviceTypeForOfflineOnline + "_PromoPlanning_DeviceTypes_Offline.xml";

            worksheet.Cells[startRowIndex, colValue++] = promoBodyEndRows;

            //added by Anwesh (22/08/2019)

            worksheet.Cells[startRowIndex, colValue++] = Convert.ToString(isDownloadEnabled);
            worksheet.Cells[startRowIndex, colValue++] = Convert.ToString(isUploadEnabled);
            worksheet.Cells[startRowIndex, colValue++] = Convert.ToString(isTCPUVDPEnabled);
            worksheet.Cells[startRowIndex, colValue++] = Convert.ToString(isBransonEnabled);
            worksheet.Cells[startRowIndex, colValue++] = Convert.ToString(FAST.userName);

            // Added for FAST Folder, if it is not available, create again
            Directory.CreateDirectory(FAST.localPathDataTable + "\\FAST");
            

            if (FAST.isDownloadEnabled)
            {
                string filePath_For_DataTableData = FAST.localPathDataTable + "\\FAST\\" + "Country" + FAST._promoDownloadCountryValueForOfflineOnline + "_Device" + FAST._promoDownloadDeviceTypeForOfflineOnline + "_PromoPlanning.xml";

                //Converting the DataTable to XML and also saving it into the XML file along with the null values
                dsPromotions.WriteXml(filePath_For_DataTableData, XmlWriteMode.WriteSchema);

                // Creating the NEW path for saving the deviceType DataSet for Offline
                string filePath_For_DataTableData_DeviceTypes = FAST.localPathDataTable + "\\FAST\\" + "Country" + FAST._promoDownloadCountryValueForOfflineOnline + "_Device" + FAST._promoDownloadDeviceTypeForOfflineOnline + "_PromoPlanning_DeviceTypes_Offline.xml";

                dsDeviceTypes.WriteXml(filePath_For_DataTableData_DeviceTypes, XmlWriteMode.WriteSchema);
            }
            else
            {
                string filePath_For_DataTableData = FAST.localPathDataTable + "\\FAST\\" + "Country" + FAST._promoCountryValue + "_Device" + FAST._promoDeviceTypeValue + "_PromoPlanning.xml";

                //Converting the DataTable to XML and also saving it into the XML file along with the null values
                //dsPromotions.WriteXml(filePath_For_DataTableData, XmlWriteMode.WriteSchema);

                // Creating the NEW path for saving the deviceType DataSet for Offline
                string filePath_For_DataTableData_DeviceTypes = FAST.localPathDataTable + "\\FAST\\" + "Country" + FAST._promoCountryValue + "_Device" + FAST._promoDeviceTypeValue + "_PromoPlanning_DeviceTypes_Offline.xml";
            }
        }

        public static void promoUploadOnline()
        {

            Worksheet worksheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[clsInformation.referencePromo]);
            if (worksheet == null)
                throw new Exception("Worksheet can not be empty when building sheet body");

            int startRowIndex = 2, colValue = 1;

            FAST._txtProcess = Convert.ToString(worksheet.Cells[startRowIndex, colValue++].Value);
            FAST._valueProcess = Convert.ToString(worksheet.Cells[startRowIndex, colValue++].Value);
            FAST._promoCountryValue = Convert.ToString(worksheet.Cells[startRowIndex, colValue++].Value);
            FAST._promoDeviceTypeValue = Convert.ToString(worksheet.Cells[startRowIndex, colValue++].Value);
            FAST._promoDownloadCountryValueForOfflineOnline = Convert.ToString(worksheet.Cells[startRowIndex, colValue++].Value);
            FAST._promoDownloadDeviceTypeForOfflineOnline = Convert.ToString(worksheet.Cells[startRowIndex, colValue++].Value);
            FAST._saveDataFile = Convert.ToString(worksheet.Cells[startRowIndex, colValue++].Value);
            FAST._promoCountryLabel = Convert.ToString(worksheet.Cells[startRowIndex, colValue++].Value);
            FAST._promoDeviceLabel = Convert.ToString(worksheet.Cells[startRowIndex, colValue++].Value);
            FAST._promoCountryRefreshTCPUVDPLable = Convert.ToString(worksheet.Cells[startRowIndex, colValue++].Value);
            FAST._promodeviceRefreshTCPUVDPLabel = Convert.ToString(worksheet.Cells[startRowIndex, colValue++].Value);
            FAST._promoCountryRefreshTCPUVDPValue = Convert.ToString(worksheet.Cells[startRowIndex, colValue++].Value);
            FAST._promodeviceRefreshTCPUVDPValue = Convert.ToString(worksheet.Cells[startRowIndex, colValue++].Value);
            FAST._promoCountryBransonRefreshLable = Convert.ToString(worksheet.Cells[startRowIndex, colValue++].Value);
            FAST._promodeviceBransonRefreshLabel = Convert.ToString(worksheet.Cells[startRowIndex, colValue++].Value);
            FAST._promoCountryBransonRefreshValue = Convert.ToString(worksheet.Cells[startRowIndex, colValue++].Value);
            FAST._promodeviceBransonRefreshValue = Convert.ToString(worksheet.Cells[startRowIndex, colValue++].Value);
            FAST.isDownloadEnabled = Convert.ToBoolean(worksheet.Cells[2, 20].Value);

            //added by Anwesh 26/08/2019
            //FAST.isUploadEnabled = Convert.ToBoolean(worksheet.Cells[2, 21].Value);
            FAST.isTCPUVDPEnabled = Convert.ToBoolean(worksheet.Cells[2, 22].Value);
            FAST.isBransonEnabled = Convert.ToBoolean(worksheet.Cells[2, 23].Value);
            FAST.userName = Convert.ToString(System.Security.Principal.WindowsIdentity.GetCurrent().Name).Split('\\')[1];
            //FAST.userName = Convert.ToString(worksheet.Cells[2, 24].Value);

            //if (FAST.userName == null)
            //{
            //    FAST.userName = Convert.ToString(System.Security.Principal.WindowsIdentity.GetCurrent().Name).Split('\\')[1];
            //}

            string deviceTypeDataXmlFileSave = Convert.ToString(worksheet.Cells[startRowIndex, colValue++].Value); // Added by Nihar on 11/30/2017
            promoBodyEndRows = Convert.ToInt16(worksheet.Cells[startRowIndex, colValue++].Value);


            string filePathforXMLData = FAST.localPathDataTable + "\\FAST\\" + FAST._saveDataFile;

            if (FAST.isDownloadEnabled)
            {
                if (filePathforXMLData != FAST.localPathDataTable)
                {
                    // Creating a dataset object here and Reading the saved xml from the xml file
                    dsPromotions = new DataSet();
                    dsPromotions.ReadXml(filePathforXMLData);
                }
            }


        }


        public static bool promotionsBeforeClose()
        {
            try
            {
                ExcelTool.Workbook excelWorkbook = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook);

                List<string> sheetNames = new List<string>();
                foreach (Excel.Worksheet sheet in excelWorkbook.Sheets)
                {
                    sheetNames.Add(sheet.Name);
                }

                if (sheetNames.Contains(clsInformation.PROMO_INPUT_TOOL))
                {
                    Worksheet worksheetPromo = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[clsInformation.PROMO_INPUT_TOOL]);

                    if (worksheetPromo.UsedRange.Rows.Count <= 1)
                        return false;

                    Excel.Range removeDropdownsFormRange = worksheetPromo.get_Range("E" + Convert.ToString(12), "H" + promoBodyEndRows) as Excel.Range;
                    removeDropdownsFormRange.Validation.Delete();
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                FAST.errorLog(ex.Message, "Promotion_Planning_Excel_promotionsBeforeClose");
                FAST.handleAlerts(ex.Message);
                FAST._IsPromotionErrorHit = true;
                return false;
            }

        }

        public static void promotionsOnOpen()
        {
            Worksheet worksheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[clsInformation.referencePromo]);
            if (worksheet == null)
                throw new Exception("Worksheet can not be empty when building sheet body");

            int startRowIndex = 2, colValue = 18;

            string deviceTypeDataXmlFileSave = Convert.ToString(worksheet.Cells[startRowIndex, colValue++].Value); // Added by Nihar on 11/30/2017
            promoBodyEndRows = Convert.ToInt16(worksheet.Cells[startRowIndex, colValue++].Value);


            // Added for Device Types XML
            string filePathforXMLDataForDeviceTypes = FAST.localPathDataTable + "\\FAST\\" + deviceTypeDataXmlFileSave;

            if (filePathforXMLDataForDeviceTypes != FAST.localPathDataTable)
            {
                // Creating a dataset object here and Reading the saved xml from the xml file
                dsDeviceTypes = new DataSet();

                if (FAST.isDownloadEnabled)
                    dsDeviceTypes.ReadXml(filePathforXMLDataForDeviceTypes);

                ExcelTool.Workbook excelWorkbook = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook);

                List<string> sheetNames = new List<string>();
                foreach (Excel.Worksheet sheet in excelWorkbook.Sheets)
                {
                    sheetNames.Add(sheet.Name);
                }

                // Worksheet worksheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[sheetNames[1]]);

                //Worksheet worksheetPromo = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[clsInformation.PROMO_INPUT_TOOL]);

                Worksheet worksheetPromo = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[sheetNames[1]]);

                unProtect(worksheetPromo);
                if (sheetNames.Contains(clsInformation.PROMO_INPUT_TOOL) && FAST._promoInputTool != null)
                    addDropdownListToColumns(worksheetPromo, 2, 12, promoBodyEndRows, "C", "E", "F", "G", "H");
                protectSheet(worksheetPromo);

            }
        }

        #endregion

        #region Refresh Branson Data
        public void refreshBransonData(string aliasId, string viewId, string countryId, string deviceTypeId, string view, string countryLabel, string deviceTypeLabel)
        {
            try
            {
                FAST.updateEvents(false);

                DataSet dsBransonData = new DataSet();

                dsBransonData = FASTWebServiceAdapter.refreshBransonDataForPromoInputTool(aliasId, viewId, countryId, deviceTypeId, view);

                if (dsBransonData.Tables.Count > 0)
                {
                    dtBransonPromotionsRows = dsBransonData.Tables["BransonPromotionsRows"];

                }
                else
                {
                    MessageBox.Show("Required Data is missing....");
                    return;
                }


                generateBransonPromotionsSheet(countryLabel, deviceTypeLabel, aliasId);


                // Selecting the PromoTool Sheet Once the work is done
                //Worksheet worksheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[clsInformation.PROMO_INPUT_TOOL]);

                //if (worksheet != null)
                //    worksheet.Select();
            }
            catch (Exception ex)
            {
                FAST.errorLog(ex.Message, "Promotion_Planning_Excel_promotionsView");
                FAST.handleAlerts(ex.Message);
                FAST._IsPromotionErrorHit = true;
            }
            finally
            {
                FAST.updateEvents(true);
            }
        }
        #endregion

    }

    public static class StringExtensions
    {
        public static string ReplaceNewLine(this String source)
        {
            return source.Replace("\r", "").Replace("\n", "");


        }

        public static string RemoveSpace(this String source)
        {
            return source.Replace(" ", "");

        }

        public static string ReplaceEmptyWithUnderScore(this String source)
        {
            return source.Replace(" ", "_");

        }



    }


}



