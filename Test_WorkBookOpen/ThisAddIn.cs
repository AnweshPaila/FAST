using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Test_WorkBookOpen.Classes;
using Microsoft.Office.Tools.Excel;
using System.Collections.Generic;
using ExcelTool = Microsoft.Office.Tools.Excel;

namespace Test_WorkBookOpen
{
    public partial class ThisAddIn
    {

        #region Startup Event
        /// <summary>
        /// This is the First Method that will be called as our application starts running.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// 

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

        }
        #endregion

        #region WorkBook Open Event
        /// <summary>
        /// Application_WorkbookOpen Event will be called whenever a saved Workbook is Opened
        /// </summary>
        /// <param name="Wb">The Activeworkbook information is passed as parameter</param>

        void Application_WorkbookOpen(Excel.Workbook Wb)
        {
            try
            {
                ExcelTool.Workbook excelWorkbook = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook);

                List<string> sheetNames = new List<string>();
                foreach (Excel.Worksheet sheet in excelWorkbook.Sheets)
                {
                    sheetNames.Add(sheet.Name);
                }

                //if (sheetNames.Contains(clsInformation.PROMO_INPUT_TOOL))
                //{

                //}
                //Worksheet worksheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[clsInformation.PROMO_INPUT_TOOL]);

                Worksheet worksheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[sheetNames[1]]);

                if (worksheet != null)
                {
                    int range = worksheet.Rows.Count; ;

                    if (range > 11)
                    {
                        ClsPromotions.promoUploadOnline();
                        FAST._verifyDownloadForUpload = true;
                    }
                }

                FAST.updateControl();
            }

            catch { }

            
        }
        #endregion

        #region Workbook Before Save Event

        /// <summary>
        /// Before Saving the Excel Workbook, this event will be called to save the modifications done by the User.
        /// </summary>
        /// <param name="Wb">Workbook information will be sent as a parameter</param>
        /// <param name="SaveAsUI"></param>
        /// <param name="Cancel"></param>
        private void Application_WorkbookBeforeSave(Excel.Workbook Wb, bool SaveAsUI, ref bool Cancel)
        {
            try
            {            
                
                if (FAST._verifyDownloadForUpload == true)
                {
                    FAST.savingRequiredDataForOffline();                    
                }
                //else if (!FAST.isDownloadEnabled)
                //{
                //    FAST.savingRequiredDataForOffline();
                //}
            }
            catch (Exception ex)
            {
                FAST.errorLog(ex.Message, "Add-in-Application_WorkbookBeforeSave");
            }
        }
        #endregion

        #region Shutdown Event
        /// <summary>
        /// When the Excel is Closed, at thay time this event is called
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// 
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            if (FAST._txtProcess == clsInformation.promotionsView)
            {
                if (!ClsPromotions.promotionsBeforeClose())
                    return;

                bool isDirty = Globals.ThisAddIn.Application.ActiveWorkbook.Saved;

                if (!isDirty)
                    return;


                Globals.ThisAddIn.Application.ActiveWorkbook.Save();
                FAST.updateEvents(true);
                Globals.ThisAddIn.Application.ActiveWorkbook.Close();
            }
        }
        #endregion

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += new System.EventHandler(ThisAddIn_Startup);
            Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
            Application.WorkbookBeforeSave += new Excel.AppEvents_WorkbookBeforeSaveEventHandler(Application_WorkbookBeforeSave);
            Application.WorkbookActivate += new Excel.AppEvents_WorkbookActivateEventHandler(Application_WorkbookOpen);
            
        }
        #endregion

    }
}
