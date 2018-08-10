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
    public class GetTrendData : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            this.GetTrendInfos(listSheet);
            return true;
        }

        private ExcelWriter GetTrendInfoExcelWriter(string destFilePath)
        {

            Dictionary<string, int> columnDic = CommonUtil.InitStringIndexDic(new string[]{ 
                    "Company_Name", 
                    "Page_Company_Name",
                    "EmployerId",
                    "2018/7/29",
                    "2018/6/24",
                    "2018/5/27",
                    "2018/4/29",
                    "2018/3/25",
                    "2018/2/25",
                    "2018/1/28",
                    "2017/12/31",
                    "2017/11/26",
                    "2017/10/29",
                    "2017/9/24",
                    "2017/8/27",
                    "2017/7/30",
                    "2017/6/25",
                    "2017/5/28",
                    "2017/4/30",
                    "2017/3/26",
                    "2017/2/26",
                    "2017/1/29",
                    "2016/12/18",
                    "2016/11/20",
                    "2016/10/23",
                    "2016/9/18",
                    "2016/8/21"
                    });
            Dictionary<string, string> columnFormats = new Dictionary<string, string>();
            columnFormats.Add("2018/7/29", "#0.00000");
            columnFormats.Add("2018/6/24", "#0.00000");
            columnFormats.Add("2018/5/27", "#0.00000");
            columnFormats.Add("2018/4/29", "#0.00000");
            columnFormats.Add("2018/3/25", "#0.00000");
            columnFormats.Add("2018/2/25", "#0.00000");
            columnFormats.Add("2018/1/28", "#0.00000");
            columnFormats.Add("2017/12/31", "#0.00000");
            columnFormats.Add("2017/11/26", "#0.00000");
            columnFormats.Add("2017/10/29", "#0.00000");
            columnFormats.Add("2017/9/24", "#0.00000");
            columnFormats.Add("2017/8/27", "#0.00000");
            columnFormats.Add("2017/7/30", "#0.00000");
            columnFormats.Add("2017/6/25", "#0.00000");
            columnFormats.Add("2017/5/28", "#0.00000");
            columnFormats.Add("2017/4/30", "#0.00000");
            columnFormats.Add("2017/3/26", "#0.00000");
            columnFormats.Add("2017/2/26", "#0.00000");
            columnFormats.Add("2017/1/29", "#0.00000");
            columnFormats.Add("2016/12/18", "#0.00000");
            columnFormats.Add("2016/11/20", "#0.00000");
            columnFormats.Add("2016/10/23", "#0.00000");
            columnFormats.Add("2016/9/18", "#0.00000");
            columnFormats.Add("2016/8/21", "#0.00000");

            ExcelWriter ew = new ExcelWriter(destFilePath, "List", columnDic, columnFormats);
            return ew;
        }

        private void GetTrendInfos(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            string resultFilePath = Path.Combine(exportDir, "GlassDoor_TrendDetail.xlsx");

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;

            Dictionary<string, string> houseDic = new Dictionary<string, string>();

            ExcelWriter resultEW = this.GetTrendInfoExcelWriter(resultFilePath);

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

                        JArray dateArray = infoJo.GetValue("dates") as JArray;
                        JArray employerRatingArray = infoJo.GetValue("employerRatings") as JArray;

                        Nullable<decimal> v2018_7_29 = null;
                        Nullable<decimal> v2018_6_24 = null;
                        Nullable<decimal> v2018_5_27 = null;
                        Nullable<decimal> v2018_4_29 = null;
                        Nullable<decimal> v2018_3_25 = null;
                        Nullable<decimal> v2018_2_25 = null;
                        Nullable<decimal> v2018_1_28 = null;
                        Nullable<decimal> v2017_12_31 = null;
                        Nullable<decimal> v2017_11_26 = null;
                        Nullable<decimal> v2017_10_29 = null;
                        Nullable<decimal> v2017_9_24 = null;
                        Nullable<decimal> v2017_8_27 = null;
                        Nullable<decimal> v2017_7_30 = null;
                        Nullable<decimal> v2017_6_25 = null;
                        Nullable<decimal> v2017_5_28 = null;
                        Nullable<decimal> v2017_4_30 = null;
                        Nullable<decimal> v2017_3_26 = null;
                        Nullable<decimal> v2017_2_26 = null;
                        Nullable<decimal> v2017_1_29 = null;
                        Nullable<decimal> v2016_12_18 = null;
                        Nullable<decimal> v2016_11_20 = null;
                        Nullable<decimal> v2016_10_23 = null;
                        Nullable<decimal> v2016_9_18 = null;
                        Nullable<decimal> v2016_8_21 = null;

                        for (int j = 0; j < dateArray.Count; j++)
                        {
                            string date = dateArray[j].ToString();
                            string employerRatingStr = employerRatingArray[j].ToString();
                            Nullable<decimal> employerRating = employerRatingStr.Length == 0 ? null : (Nullable<decimal>)decimal.Parse(employerRatingStr);
                            switch (date)
                            {
                                case "2018/7/29":
                                    v2018_7_29 = employerRating; break;
                                case "2018/6/24":
                                    v2018_6_24 = employerRating; break;
                                case "2018/5/27":
                                    v2018_5_27 = employerRating; break;
                                case "2018/4/29":
                                    v2018_4_29 = employerRating; break;
                                case "2018/3/25":
                                    v2018_3_25 = employerRating; break;
                                case "2018/2/25":
                                    v2018_2_25 = employerRating; break;
                                case "2018/1/28":
                                    v2018_1_28 = employerRating; break;
                                case "2017/12/31":
                                    v2017_12_31 = employerRating; break;
                                case "2017/11/26":
                                    v2017_11_26 = employerRating; break;
                                case "2017/10/29":
                                    v2017_10_29 = employerRating; break;
                                case "2017/9/24":
                                    v2017_9_24 = employerRating; break;
                                case "2017/8/27":
                                    v2017_8_27 = employerRating; break;
                                case "2017/7/30":
                                    v2017_7_30 = employerRating; break;
                                case "2017/6/25":
                                    v2017_6_25 = employerRating; break;
                                case "2017/5/28":
                                    v2017_5_28 = employerRating; break;
                                case "2017/4/30":
                                    v2017_4_30 = employerRating; break;
                                case "2017/3/26":
                                    v2017_3_26 = employerRating; break;
                                case "2017/2/26":
                                    v2017_2_26 = employerRating; break;
                                case "2017/1/29":
                                    v2017_1_29 = employerRating; break;
                                case "2016/12/18":
                                    v2016_12_18 = employerRating; break;
                                case "2016/11/20":
                                    v2016_11_20 = employerRating; break;
                                case "2016/10/23":
                                    v2016_10_23 = employerRating; break;
                                case "2016/9/18":
                                    v2016_9_18 = employerRating; break;
                                case "2016/8/21":
                                    v2016_8_21 = employerRating; break;
                            }
                        }

                        Dictionary<string, object> resultRow = new Dictionary<string, object>();
                        resultRow.Add("Company_Name", companyName);
                        resultRow.Add("Page_Company_Name", pageCompanyName);
                        resultRow.Add("EmployerId", employerId);

                        resultRow.Add("2018/7/29", v2018_7_29.HasValue ? (object)v2018_7_29 : null);
                        resultRow.Add("2018/6/24", v2018_6_24.HasValue ? (object)v2018_6_24 : null);
                        resultRow.Add("2018/5/27", v2018_5_27.HasValue ? (object)v2018_5_27 : null);
                        resultRow.Add("2018/4/29", v2018_4_29.HasValue ? (object)v2018_4_29 : null);
                        resultRow.Add("2018/3/25", v2018_3_25.HasValue ? (object)v2018_3_25 : null);
                        resultRow.Add("2018/2/25", v2018_2_25.HasValue ? (object)v2018_2_25 : null);
                        resultRow.Add("2018/1/28", v2018_1_28.HasValue ? (object)v2018_1_28 : null);
                        resultRow.Add("2017/12/31", v2017_12_31.HasValue ? (object)v2017_12_31 : null);
                        resultRow.Add("2017/11/26", v2017_11_26.HasValue ? (object)v2017_11_26 : null);
                        resultRow.Add("2017/10/29", v2017_10_29.HasValue ? (object)v2017_10_29 : null);
                        resultRow.Add("2017/9/24", v2017_9_24.HasValue ? (object)v2017_9_24 : null);
                        resultRow.Add("2017/8/27", v2017_8_27.HasValue ? (object)v2017_8_27 : null);
                        resultRow.Add("2017/7/30", v2017_7_30.HasValue ? (object)v2017_7_30 : null);
                        resultRow.Add("2017/6/25", v2017_6_25.HasValue ? (object)v2017_6_25 : null);
                        resultRow.Add("2017/5/28", v2017_5_28.HasValue ? (object)v2017_5_28 : null);
                        resultRow.Add("2017/4/30", v2017_4_30.HasValue ? (object)v2017_4_30 : null);
                        resultRow.Add("2017/3/26", v2017_3_26.HasValue ? (object)v2017_3_26 : null);
                        resultRow.Add("2017/2/26", v2017_2_26.HasValue ? (object)v2017_2_26 : null);
                        resultRow.Add("2017/1/29", v2017_1_29.HasValue ? (object)v2017_1_29 : null);
                        resultRow.Add("2016/12/18", v2016_12_18.HasValue ? (object)v2016_12_18 : null);
                        resultRow.Add("2016/11/20", v2016_11_20.HasValue ? (object)v2016_11_20 : null);
                        resultRow.Add("2016/10/23", v2016_10_23.HasValue ? (object)v2016_10_23 : null);
                        resultRow.Add("2016/9/18", v2016_9_18.HasValue ? (object)v2016_9_18 : null);
                        resultRow.Add("2016/8/21", v2016_8_21.HasValue ? (object)v2016_8_21 : null);
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