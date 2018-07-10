using HtmlAgilityPack;
using NetDataAccess.Base.Common;
using NetDataAccess.Base.Config;
using NetDataAccess.Base.DB;
using NetDataAccess.Base.UI;
using NetDataAccess.Base.Writer;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Web;

namespace NetDataAccess.Extended.Xingzheng.TongJiYongQuHuaDaiMa
{
    public class Level3 : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            this.GetCities(listSheet);
            return true;
        }

        private void GetCities(IListSheet listSheet)
        {
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            ExcelWriter ew = this.GetExcelWriter();

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

                Uri uri = new Uri(detailUrl); 
                string queryString = uri.Query;
                string baseUrl = detailUrl.Substring(0, detailUrl.Length - queryString.Length);
                baseUrl = baseUrl.Substring(0, baseUrl.Length - uri.Segments[uri.Segments.Length - 1].Length);

                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
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
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
            }

            ew.SaveToDisk();
        }


        private ExcelWriter GetExcelWriter()
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
                    "name"});

            string filePath = Path.Combine(exportDir, "统计用区划代码_Level4.xlsx");
            ExcelWriter ew = new ExcelWriter(filePath, "List", columnDic);
            return ew;
        }
    }
}
