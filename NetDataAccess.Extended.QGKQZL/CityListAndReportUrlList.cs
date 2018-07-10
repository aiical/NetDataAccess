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
    /// CityListAndReportUrlList
    /// </summary>
    public class CityListAndReportUrlList : CustomProgramBase
    {
        public bool Run(string parameters, IListSheet listSheet)
        {
            return this.GenerateCityListAndReportList(parameters, listSheet);
        }
        private bool GenerateCityListAndReportList(string parameters, IListSheet listSheet)
        {
            string[] times = parameters.Split(new string[]{","}, StringSplitOptions.RemoveEmptyEntries);
            int fromYear = int.Parse(times[0]);
            int fromMonth = int.Parse(times[1]);
            int toYear = int.Parse(times[2]);
            int toMonth = int.Parse(times[3]);

            bool succeed = true;
            string exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> cityListColumnDic = new Dictionary<string, int>();
            cityListColumnDic.Add("province", 0);
            cityListColumnDic.Add("city", 1);
            cityListColumnDic.Add("cityCode", 2);
            string cityListPath = Path.Combine(exportDir, this.RunPage.Project.Name + "_List.xlsx");
            ExcelWriter cityListEW = new ExcelWriter(cityListPath, "List", cityListColumnDic);

            Dictionary<string, int> cityReportListColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                "detailPageUrl",
                "detailPageName",
                "cookie", 
                "grabStatus", 
                "giveUpGrab",
                "cityCode",
                "cityName", 
                "year",
                "month"});
            string cityReportListFilePath = Path.Combine(exportDir, "城市空气报告列表.xlsx");
            ExcelWriter cityReportListEW = new ExcelWriter(cityReportListFilePath, "List", cityReportListColumnDic);
            int cityReportPageIndex = 1;
              
            int detailUrlColumnIndex = this.RunPage.ColumnNameToIndex["detailPageUrl"];

            List<string> allCityCodes = new List<string>();

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string,string> row = listSheet.GetRow(i);
                string giveUp = row["giveUpGrab"];
                if (giveUp != "是")
                {
                    string detailUrl = row["detailPageUrl"];
                    string localFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir);
                    TextReader tr = null;

                    try
                    {
                        tr = new StreamReader(localFilePath, Encoding.GetEncoding(((Proj_Detail_SingleLine)this.RunPage.Project.DetailGrabInfoObject).Encoding));
                        string webPageHtml = tr.ReadToEnd();

                        HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                        htmlDoc.LoadHtml(webPageHtml);

                        HtmlNodeCollection listDivNodeList = htmlDoc.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/dl");
                        foreach (HtmlNode listDivNode in listDivNodeList)
                        {
                            HtmlNode provinceNode = listDivNode.SelectSingleNode("./dt");
                            string province = provinceNode.InnerText;

                            HtmlNodeCollection cityNodeList = listDivNode.SelectNodes("./dd/a");
                            foreach (HtmlNode cityNode in cityNodeList)
                            {
                                string city = cityNode.InnerText.Trim();
                                string cityUrl = cityNode.GetAttributeValue("href", "");
                                int startIndex = cityUrl.IndexOf("/aqi/") + 5;
                                int endIndex = cityUrl.IndexOf(".html");
                                string cityCode = cityUrl.Substring(startIndex, endIndex - startIndex).Trim();
                                Dictionary<string, string> cityInfo = new Dictionary<string, string>();
                                cityInfo.Add("province", province);
                                cityInfo.Add("city", city);
                                cityInfo.Add("cityCode", cityCode);
                                cityListEW.AddRow(cityInfo);

                                if (!allCityCodes.Contains(cityCode))
                                {
                                    allCityCodes.Add(cityCode);

                                    for (int year = fromYear; year <= toYear; year++)
                                    {
                                        for (int month = 1; month <= 12; month++)
                                        {
                                            if ((year == fromYear && month < fromMonth) || (year == toYear && month > toMonth))
                                            {
                                                continue;
                                            }
                                            else
                                            {
                                                string monthStr = month.ToString("00");
                                                string urlName = cityCode + "-" + year.ToString() + monthStr;
                                                string url = "http://www.tianqihoubao.com/aqi/" + urlName + ".html";
                                                Dictionary<string, string> cityReport = new Dictionary<string, string>();
                                                cityReport.Add("detailPageUrl", url);
                                                cityReport.Add("detailPageName", urlName);
                                                cityReport.Add("cityCode", cityCode);
                                                cityReport.Add("cityName", city);
                                                cityReport.Add("year", year.ToString());
                                                cityReport.Add("month", monthStr);
                                                cityReportListEW.AddRow(cityReport);
                                                cityReportPageIndex++;
                                            }
                                        }
                                    }
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
            }
            cityListEW.SaveToDisk();
            cityReportListEW.SaveToDisk(); 
            return succeed;
        }
         
    }
}