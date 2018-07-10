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

namespace NetDataAccess.Extended.QGKQZL
{
    /// <summary>
    /// CityReport
    /// </summary>
    public class CityReport : CustomProgramBase
    {
        public bool Run(string parameters, IListSheet listSheet)
        {
            return this.GenerateCityReport(listSheet);
        }
        private bool GenerateCityReport(IListSheet listSheet)
        {
            bool succeed = true;
            string exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> cityReportColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                        "cityCode",
                        "city",
                        "日期", 
                        "AQI指数", 
                        "质量等级", 
                        "当天AQI排名", 
                        "PM2.5",
                        "PM10", 
                        "Co", 
                        "No2", 
                        "So2",
                        "O3"});
            string cityReportPath = Path.Combine(exportDir, this.RunPage.Project.Name + "_List.xlsx");
            Dictionary<string, string> columnFormats = new Dictionary<string, string>();
            columnFormats.Add("日期", "yyyy-m-d");
            columnFormats.Add("AQI指数", "#0");
            columnFormats.Add("当天AQI排名", "#0");
            columnFormats.Add("PM2.5", "#0");
            columnFormats.Add("PM10", "#0");
            columnFormats.Add("Co", "#0.00");
            columnFormats.Add("No2", "#0");
            columnFormats.Add("So2", "#0");
            columnFormats.Add("O3", "#0");
            ExcelWriter cityReportEW = new ExcelWriter(cityReportPath, "List", cityReportColumnDic, columnFormats);


            int detailUrlColumnIndex = this.RunPage.ColumnNameToIndex["detailPageUrl"];
            Dictionary<string, string> codeDateToNull = new Dictionary<string, string>();
            string sourceDateFormat = "yyyy-MM-dd";

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string,string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"];
                string cityCode = row["cityCode"];
                string city = row["cityName"];
                string localFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir);
                TextReader tr = null;

                try
                {
                    tr = new StreamReader(localFilePath, Encoding.GetEncoding(((Proj_Detail_SingleLine)this.RunPage.Project.DetailGrabInfoObject).Encoding));
                    string webPageHtml = tr.ReadToEnd();

                    HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                    htmlDoc.LoadHtml(webPageHtml);

                    HtmlNodeCollection listDivNodeList = htmlDoc.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[3]/table[1]/tr");
                    if (listDivNodeList.Count > 1)
                    {
                        Dictionary<int, string> cityReportColumnIndexDic = new Dictionary<int, string>();
                        HtmlNodeCollection nameNodes = listDivNodeList[0].SelectNodes("td");
                        for (int j = 0; j < nameNodes.Count; j++)
                        {
                            HtmlNode nameNode = nameNodes[j];
                            string name = nameNode.InnerText.Trim();
                            cityReportColumnIndexDic.Add(j, name);
                        }
                        for (int j = 1; j < listDivNodeList.Count; j++)
                        {
                            HtmlNode listDivNode = listDivNodeList[j];

                            HtmlNodeCollection vNodeList = listDivNode.SelectNodes("./td");
                            Dictionary<string, object> reportInfo = new Dictionary<string, object>();
                            reportInfo.Add("cityCode", cityCode);
                            reportInfo.Add("city", city);

                            for (int k = 0; k < nameNodes.Count; k++)
                            {
                                HtmlNode vNode = vNodeList[k];
                                string value = vNode.InnerText.Trim(); 
                                string columName = cityReportColumnIndexDic[k];
                                switch (columName)
                                {
                                    case "日期":
                                        DateTime dt = DateTime.ParseExact(value, sourceDateFormat, System.Globalization.CultureInfo.CurrentCulture);
                                        reportInfo.Add(columName, dt);
                                        break;
                                    case "AQI指数":
                                    case "当天AQI排名":
                                    case "PM2.5":
                                    case "PM10":
                                    case "Co":
                                    case "No2":
                                    case "So2": 
                                        reportInfo.Add(columName, decimal.Parse(value)); 
                                        break;
                                    default:
                                        reportInfo.Add(columName, value);
                                        break;

                                }
                            }
                            string codeDate = cityCode + "_" + ((DateTime)reportInfo["日期"]).ToString("yyyy-MM-dd");
                            if (!codeDateToNull.ContainsKey(codeDate))
                            {
                                cityReportEW.AddRow(reportInfo);
                                codeDateToNull.Add(codeDate, null);
                            }

                        }
                    }
                }
                catch (Exception ex)
                {
                    if (tr != null)
                    {
                        tr.Dispose();
                        tr = null;
                    }
                    this.RunPage.InvokeAppendLogText("读取出错. " + ex.Message + " LocalPath = " + localFilePath, LogLevelType.Error, true);
                }
            }
            cityReportEW.SaveToDisk(); 
            return succeed;
        }
         
    }
}