using System;
using System.Collections.Generic;
using System.Text;
using NetDataAccess.Base.DLL;
using NetDataAccess.Base.Config;
using System.Threading;
using System.Windows.Forms;
using mshtml;
using NetDataAccess.Base.Definition;
using System.IO;
using NetDataAccess.Base.Common;
using NPOI.SS.UserModel;
using NetDataAccess.Base.EnumTypes;
using NetDataAccess.Base.UI;
using Newtonsoft.Json.Linq;
using HtmlAgilityPack;
using NetDataAccess.Base.Writer;
using NetDataAccess.Base.DB;
using NetDataAccess.Base.CsvHelper;
using NetDataAccess.Base.Reader;

namespace NetDataAccess.Extended.GlassDoor
{
    public class GetOverallDistributionData : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            this.GetOverallDistributionInfos(listSheet);
            return true;
        }

        private ExcelWriter GetOverallDistributionInfoExcelWriter(string destFilePath)
        {

            Dictionary<string, int> columnDic = CommonUtil.InitStringIndexDic(new string[]{ 
                    "Company_Name", 
                    "Page_Company_Name",
                    "EmployerId",
                    "1 Star",
                    "2 Stars",
                    "3 Stars",
                    "4 Stars",
                    "5 Stars"
            });
            Dictionary<string, string> columnFormats = new Dictionary<string, string>();
            columnFormats.Add("1 Star", "#0");
            columnFormats.Add("2 Stars", "#0");
            columnFormats.Add("3 Stars", "#0");
            columnFormats.Add("4 Stars", "#0");
            columnFormats.Add("5 Stars", "#0");

            ExcelWriter ew = new ExcelWriter(destFilePath, "List", columnDic, columnFormats);
            return ew;
        }

        private void GetOverallDistributionInfos(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            string resultFilePath = Path.Combine(exportDir, "GlassDoor_OverallDistributionDetail.xlsx");

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;

            Dictionary<string, string> houseDic = new Dictionary<string, string>();

            ExcelWriter resultEW = this.GetOverallDistributionInfoExcelWriter(resultFilePath);

            for (int i = 0; i < listSheet.RowCount; i++)
            {

                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string url = row[detailPageUrlColumnName];
                    string companyName = row["Company_Name"];
                    string pageCompanyName = row["Page_Company_Name"];
                    string employerId = row["EmployerId"];

                    string pageFilePath = this.RunPage.GetFilePath(url, pageSourceDir);
                    string jsonText = FileHelper.GetTextFromFile(pageFilePath);


                    try
                    {
                        JObject infoJo = JObject.Parse(jsonText);

                        JArray labelArray = infoJo.GetValue("labels") as JArray;
                        JArray valueArray = infoJo.GetValue("values") as JArray;

                        Nullable<decimal> star1 = null;
                        Nullable<decimal> star2 = null;
                        Nullable<decimal> star3 = null;
                        Nullable<decimal> star4 = null;
                        Nullable<decimal> star5 = null;

                        for (int j = 0; j < labelArray.Count; j++)
                        {
                            string label  = labelArray[j].ToString();
                            string valueStr = valueArray[j].ToString();
                            Nullable<decimal> value = valueStr.Length == 0 ? null : (Nullable<decimal>)decimal.Parse(valueStr);

                            switch (label)
                            {
                                case "1 Star":
                                    star1 = value;
                                    break;
                                case "2 Stars":
                                    star2 = value;
                                    break;
                                case "3 Stars":
                                    star3 = value;
                                    break;
                                case "4 Stars":
                                    star4 = value;
                                    break;
                                case "5 Stars":
                                    star5 = value;
                                    break;
                            }

                        }

                        Dictionary<string, object> resultRow = new Dictionary<string, object>();
                        resultRow.Add("Company_Name", companyName);
                        resultRow.Add("Page_Company_Name", pageCompanyName);
                        resultRow.Add("EmployerId", employerId);
                        resultRow.Add("1 Star", star1.HasValue ? (object)star1 : null);
                        resultRow.Add("2 Stars", star2.HasValue ? (object)star2 : null);
                        resultRow.Add("3 Stars", star3.HasValue ? (object)star3 : null);
                        resultRow.Add("4 Stars", star4.HasValue ? (object)star4 : null);
                        resultRow.Add("5 Stars", star5.HasValue ? (object)star5 : null);
                        resultEW.AddRow(resultRow);
                    }
                    catch (Exception ex)
                    {
                        this.RunPage.InvokeAppendLogText(ex.Message + ", pageUrl = " + url, LogLevelType.System, true);
                        throw ex;
                    }
                }
            }

            resultEW.SaveToDisk();
        } 
    }
}