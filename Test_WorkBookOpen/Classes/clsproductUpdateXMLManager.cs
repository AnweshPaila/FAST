using System;
using System.Data;
using System.IO;
using System.Xml.Linq;
using ExcelTool = Microsoft.Office.Tools.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Text;

namespace Test_WorkBookOpen.Classes
{
    class clsproductUpdateXMLManager
    {
        #region Variable Decleration

        public static XDocument _data = null;
        public static string _localPath = null;
        public static XElement xe = null;

        //Added by Sameera
        public static StringBuilder sb;

        #endregion

        #region Constructor

        /// <summary>
        /// The constructor is called get the path of a file to save the data for Offline
        /// </summary>
        /// <param name="name"></param>

        public clsproductUpdateXMLManager(string name)
        {

            _localPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + string.Format("\\{0}.xml", name);
            clearXmlRoot();

        }
        #endregion

        #region Public Method
        /// <summary>
        /// Used to Insert the data that has been modified by the User
        /// </summary>
        /// <param name="dr">Gives the row Information</param>
        /// <param name="dc">gives the column Information</param>

        public static void convertRangeToXml()
        {
            ExcelTool.Worksheet worksheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[clsInformation.productRevenue]);


            string[] columns = FAST._dsDownloadData.Tables[1].Rows[0].ItemArray[0].ToString().Split(',');

            if (FAST._txtProcess == clsInformation.tcpuView)
                FAST._lastColumnName = clsManageSheet.getColumnName(columns.Length);// removed the last column as it is Readonly and will be used only for TCPU
            else
                FAST._lastColumnName = clsManageSheet.getColumnName(columns.Length + 1);





            Excel.Range header = worksheet.get_Range(clsManageSheet.formulaNextColumn + clsManageSheet.bodyRowStartingNumber, FAST._lastColumnName + clsManageSheet.bodyRowStartingNumber) as Excel.Range;

            int rowCount = FAST._dsDownloadData.Tables[2].Rows.Count + clsManageSheet.bodyRowStartingNumber;

            DataTable dt = FAST.makeTableFromRange(header, worksheet, rowCount);


            List<string> table = new List<string>();
            string[] months = new string[] { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" };


            foreach (Excel.Range col in header)
            {

                if (FAST._txtProcess == clsInformation.accountingView)
                {
                    string[] abc = Convert.ToString(col.Text).Split('/');
                    int a = Convert.ToInt16(abc[0]) - 1;
                    string val = months[a] + "' " + abc[2].Substring(abc[2].Length - 2);
                    table.Add(val);
                }
                else
                {
                    //table.Add(Convert.ToString(col.Text).Replace("'", "-"));
                    table.Add(Convert.ToString(col.Text).Replace("'", @"\' "));
                }
            }

            if (FAST._txtProcess == clsInformation.accountingView)
            {
                #region  Xml
                clearXmlRoot();

                foreach (DataRow dr in dt.Rows)
                {
                    foreach (var item in table)
                    {
                        _data.Element("XMLRoot").Add(new XElement("UpdatingRow",

                           new XElement("ProductLineId", dr["ProductLineId"]),
                           new XElement("ChannelId", dr["ChannelId"]),
                           new XElement("CountryId", dr["CountryId"]),
                           new XElement("ProgramId", dr["ProgramId"]),
                           new XElement("MemoryId", dr["MemoryId"]),
                           new XElement("WirelessId", dr["WirelessId"]),
                           new XElement("DTCPId", dr["DTCPId"]),
                           new XElement("CurrencyId", dr["CurrencyId"]),
                            new XElement("MonthYear", item),
                           new XElement("InputTypeValue", dr[item]),
                           new XElement("InputTemplateDataId", dr["InputTemplateDataId"])
                           ));
                    }
                }
                #endregion

                #region Commented CSV
                //added by Praveen
                //sb = new StringBuilder();
                //string ddInputvalue = null;

                //foreach (DataRow dr in dt.Rows)
                //{
                //    foreach (var item in table)
                //    {
                //        #region Commented
                //        //_data.Element("XMLRoot").Add(new XElement("UpdatingRow",

                //        //   new XElement("ProductLineId", dr["ProductLineId"]),
                //        //   new XElement("ChannelId", dr["ChannelId"]),
                //        //   new XElement("CountryId", dr["CountryId"]),
                //        //   new XElement("ProgramId", dr["ProgramId"]),
                //        //   new XElement("MemoryId", dr["MemoryId"]),
                //        //   new XElement("WirelessId", dr["WirelessId"]),
                //        //   new XElement("DTCPId", dr["DTCPId"]),
                //        //   new XElement("CurrencyId", dr["CurrencyId"]),
                //        //    new XElement("MonthYear", item),
                //        //   new XElement("InputTypeValue", dr[item]),
                //        //   new XElement("InputTemplateDataId", dr["InputTemplateDataId"])
                //        //   ));
                //        #endregion

                //        ////Added by Praveen
                //        //if (dr[item].ToString() == "")
                //        //    ddInputvalue = "null";
                //        //else
                //        //    ddInputvalue = dr[item].ToString();

                //        //sb.AppendFormat("({0}, {1}, {2},{3},{4},{5},{6},{7},{8},{9},{10}),", dr["ProductLineId"], dr["ChannelId"], dr["CountryId"], dr["ProgramId"], dr["MemoryId"], dr["WirelessId"], dr["DTCPId"], dr["CurrencyId"], "\"" + item + "\"", ddInputvalue, dr["InputTemplateDataId"]);

                //    }
                //}
                // sb = sb.Remove(sb.Length - 1, 1);
                #endregion
            }
            else if (FAST._txtProcess == clsInformation.tcpuView)
            {

                sb = new StringBuilder();
                string ddInputvalue = null;

                sb.Clear();

                foreach (DataRow dr in dt.Rows)
                {
                    foreach (var item in table)
                    {
                        string itemValue = null;

                        if (item != "Life Time Value")
                            itemValue = Convert.ToString(item).Replace(" ", string.Empty).Replace("'", "-").Replace("\\", string.Empty);
                        else
                            itemValue = item;

                        if (dr[itemValue].ToString() == "")
                            ddInputvalue = "null";
                        else
                            ddInputvalue = dr[itemValue].ToString();

                        //Added by Sameera
                        //sb.AppendFormat("({0}, {1}, {2},{3}),", dr["CurrencyId"], "\"" + item + "\"", ddInputvalue, dr["InputTemplateDataId"]);

                        sb.AppendFormat("({0}, {1}, {2}),", "\"" + item + "\"", ddInputvalue, dr["InputTemplateDataId"]);
                    }
                }
                //Added by Sameera
                sb = sb.Remove(sb.Length - 1, 1);
            }
        }


        public static string ConvertToCSV(String xmlInputString)
        {

            //string xmlInput = xmlInputString;
            //string csvOut = string.Empty;
            //string strNodeValue = null;
            //XDocument doc = XDocument.Parse(xmlInput);
            //StringBuilder sb = new StringBuilder();

            //int i = 1;


            //foreach (XElement node in doc.Descendants("UpdatingRow"))
            //{
            //	int j = 1;
            //	foreach (XElement innerNode in node.Elements())
            //	{

            //		//sb.AppendFormat("{0}, {1}, {2},{3}", doc.Element("CurrencyId"), doc.Element("MonthYear"), doc.Element("InputTypeValue"), doc.Element("InputTemplateDataId"));
            //		if (innerNode.Name == "CurrencyId" || innerNode.Name == "MonthYear" || innerNode.Name == "InputTypeValue" || innerNode.Name == "InputTemplateDataId")
            //		{
            //			//Month year ""
            //			if (innerNode.Name == "MonthYear")
            //				strNodeValue = "\"" + innerNode.Value.ToString() + "\"";
            //			else if (innerNode.Name == "InputTypeValue")
            //			{
            //				strNodeValue = innerNode.Value.ToString();
            //				if (strNodeValue == "")
            //					strNodeValue = "null";
            //			}
            //			else
            //				strNodeValue = innerNode.Value.ToString();


            //			if (j < 11)
            //				sb.AppendFormat("{0}," , strNodeValue);
            //			else
            //				sb.AppendFormat("{0}", strNodeValue);
            //		}

            //		j++;
            //	}

            //	if (i < doc.Descendants("UpdatingRow").Count())
            //		sb.Append("),(");
            //	i++;
            //}

            //	return sb.ToString();

            return xmlInputString;

        }


        /// <summary>
        /// Used to clear any data is available in XML Root
        /// </summary>

        public static void clearXmlRoot()
        {
            _data = new XDocument();

            XElement xePFED = new XElement("XMLRoot");
            _data.Add(xePFED);

        }

        public static void deleteFile()
        {
            if (File.Exists(_localPath))
            {
                // Deleting the edit xml files from hidden path of special folder
                File.Delete(_localPath);
            }
        }
        #endregion
    }
}
