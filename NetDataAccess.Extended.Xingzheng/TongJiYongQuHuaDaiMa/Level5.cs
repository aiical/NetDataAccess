using HtmlAgilityPack;
using NetDataAccess.Base.Common;
using NetDataAccess.Base.Config;
using NetDataAccess.Base.DB;
using NetDataAccess.Base.EnumTypes;
using NetDataAccess.Base.UI;
using NetDataAccess.Base.Writer;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;

namespace NetDataAccess.Extended.Xingzheng.TongJiYongQuHuaDaiMa
{
    public class Level5 : ExternalRunWebPage
    {
        public override bool AfterGrabOneCatchException(string pageUrl, Dictionary<string, string> listRow, Exception ex)
        {
            if (ex.InnerException is GiveUpException)
            {
                this.RunPage.InvokeAppendLogText(ex.Message + ", pageUrl = " + pageUrl, LogLevelType.Error, true);
                return true;
            }
            else if (ex.InnerException is WebException)
            {
                WebException webEx = (WebException)ex.InnerException;
                if (webEx.Status == WebExceptionStatus.ProtocolError)
                {
                    HttpWebResponse webRes = (HttpWebResponse)webEx.Response;
                    if (webRes.StatusCode == HttpStatusCode.NotFound)
                    {
                        this.RunPage.InvokeAppendLogText("服务器端不存在此网页(404), pageUrl = " + pageUrl, LogLevelType.System, true);
                        return true;
                    }
                }
            }
            return false;
        }

        public override bool AfterAllGrab(IListSheet listSheet)
        {
            this.GetCities(listSheet);
            return true;
        }

        private void GetCities(IListSheet listSheet)
        {
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            CsvWriter ew = this.GetCsvWriter();

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"];
                string detailName = row["detailPageName"];
                string year = row["year"];
                string parentCode = row["code"];
                string parentName = row["name"];

                //添加父节点到下一级文件
                Dictionary<string, string> parentF2vs = new Dictionary<string, string>();
                parentF2vs.Add("detailPageUrl", detailUrl);
                parentF2vs.Add("detailPageName", detailName);
                parentF2vs.Add("year", year);
                parentF2vs.Add("code", parentCode);
                parentF2vs.Add("name", parentName);
                parentF2vs.Add("giveUpGrab", "Y");
                ew.AddRow(parentF2vs);

                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {

                    Uri uri = new Uri(detailUrl);
                    string queryString = uri.Query;
                    string baseUrl = detailUrl.Substring(0, detailUrl.Length - queryString.Length);
                    baseUrl = baseUrl.Substring(0, baseUrl.Length - uri.Segments[uri.Segments.Length - 1].Length);

                    HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i, Encoding.GetEncoding("gb2312"));
                    try
                    {

                        HtmlNodeCollection cityNodeList = htmlDoc.DocumentNode.SelectNodes("//tr[@class=\"citytr\"]");
                        if (cityNodeList != null)
                        {
                            for (int j = 0; j < cityNodeList.Count; j++)
                            {
                                HtmlNode cityNode = cityNodeList[j];
                                HtmlNodeCollection cityFieldNodeList = cityNode.SelectNodes("./td");
                                HtmlNode cityCodeNode = cityFieldNodeList[0];
                                HtmlNode cityNameNode = cityFieldNodeList[1];
                                string cityCode = cityCodeNode.InnerText.Trim();
                                string cityName = cityNameNode.InnerText.Trim();
                                HtmlNode linkNode = cityCodeNode.SelectSingleNode("./a");

                                string hrefValue = "";
                                if (linkNode != null)
                                {
                                    hrefValue = linkNode.GetAttributeValue("href", "");
                                }

                                Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                f2vs.Add("detailPageUrl", baseUrl + hrefValue);
                                f2vs.Add("detailPageName", year + "_" + cityCode);
                                f2vs.Add("year", year);
                                f2vs.Add("code", cityCode);
                                f2vs.Add("name", cityName);
                                f2vs.Add("giveUpGrab", hrefValue.Length == 0 ? "Y" : "");
                                ew.AddRow(f2vs);
                            }
                        }

                        HtmlNodeCollection townNodeList = htmlDoc.DocumentNode.SelectNodes("//tr[@class=\"towntr\"]");
                        if (townNodeList != null)
                        {
                            for (int j = 0; j < townNodeList.Count; j++)
                            {
                                HtmlNode townNode = townNodeList[j];
                                HtmlNodeCollection townFieldNodeList = townNode.SelectNodes("./td");
                                HtmlNode townCodeNode = townFieldNodeList[0];
                                HtmlNode townNameNode = townFieldNodeList[1];
                                string townCode = townCodeNode.InnerText.Trim();
                                string townName = townNameNode.InnerText.Trim();
                                HtmlNode linkNode = townCodeNode.SelectSingleNode("./a");

                                string hrefValue = "";
                                if (linkNode != null)
                                {
                                    hrefValue = linkNode.GetAttributeValue("href", "");
                                }

                                Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                f2vs.Add("detailPageUrl", baseUrl + hrefValue);
                                f2vs.Add("detailPageName", year + "_" + townCode);
                                f2vs.Add("year", year);
                                f2vs.Add("code", townCode);
                                f2vs.Add("name", townName);
                                f2vs.Add("giveUpGrab", hrefValue.Length == 0 ? "Y" : "");
                                ew.AddRow(f2vs);
                            }
                        }

                        HtmlNodeCollection countyNodeList = htmlDoc.DocumentNode.SelectNodes("//tr[@class=\"countytr\"]");
                        if (countyNodeList != null)
                        {
                            for (int j = 0; j < countyNodeList.Count; j++)
                            {
                                HtmlNode countyNode = countyNodeList[j];
                                HtmlNodeCollection countyFieldNodeList = countyNode.SelectNodes("./td");
                                HtmlNode countyCodeNode = countyFieldNodeList[0];
                                HtmlNode countyNameNode = countyFieldNodeList[1];
                                string countyCode = countyCodeNode.InnerText.Trim();
                                string countyName = countyNameNode.InnerText.Trim();
                                HtmlNode linkNode = countyCodeNode.SelectSingleNode("./a");

                                string hrefValue = "";
                                if (linkNode != null)
                                {
                                    hrefValue = linkNode.GetAttributeValue("href", "");
                                }

                                Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                f2vs.Add("detailPageUrl", baseUrl + hrefValue);
                                f2vs.Add("detailPageName", year + "_" + countyCode);
                                f2vs.Add("year", year);
                                f2vs.Add("code", countyCode);
                                f2vs.Add("name", countyName);
                                f2vs.Add("giveUpGrab", hrefValue.Length == 0 ? "Y" : "");
                                ew.AddRow(f2vs);
                            }
                        }

                        HtmlNodeCollection villageNodeList = htmlDoc.DocumentNode.SelectNodes("//tr[@class=\"villagetr\"]");
                        if (villageNodeList != null)
                        {
                            for (int j = 0; j < villageNodeList.Count; j++)
                            {
                                HtmlNode villageNode = villageNodeList[j];
                                HtmlNodeCollection villageFieldNodeList = villageNode.SelectNodes("./td");
                                HtmlNode villageCodeNode = villageFieldNodeList[0];
                                HtmlNode villageTypeNode = villageFieldNodeList[1];
                                HtmlNode villageNameNode = villageFieldNodeList[2];
                                string villageCode = villageCodeNode.InnerText.Trim();
                                string villageType = villageTypeNode.InnerText.Trim();
                                string villageName = villageNameNode.InnerText.Trim(); 

                                Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                f2vs.Add("detailPageUrl", year + "_" + villageCode);
                                f2vs.Add("detailPageName", year + "_" + villageCode);
                                f2vs.Add("year", year);
                                f2vs.Add("code", villageCode);
                                f2vs.Add("name", villageName);
                                f2vs.Add("giveUpGrab", "Y");
                                f2vs.Add("城乡分类代码", villageType);
                                ew.AddRow(f2vs);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
            }

            ew.SaveToDisk();
        }


        private CsvWriter GetCsvWriter()
        {
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> columnDic = CommonUtil.InitStringIndexDic(new string[]{
                    "detailPageUrl",
                    "detailPageName", 
                    "cookie",
                    "grabStatus", 
                    "giveUpGrab",
                    "year",
                    "code",
                    "name",
                    "城乡分类代码"});

            string filePath = Path.Combine(exportDir, "统计用区划代码.csv");
            CsvWriter ew = new CsvWriter(filePath, columnDic);
            return ew;
        }
    }
}
