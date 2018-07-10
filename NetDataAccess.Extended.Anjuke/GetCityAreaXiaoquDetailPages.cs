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
using NetDataAccess.Base.Web;
using System.Net;

namespace NetDataAccess.Extended.Anjuke
{
    public class GetCityAreaXiaoquDetailPages : ExternalRunWebPage
    {
        public override void WebRequestHtml_BeforeSendRequest(string pageUrl, Dictionary<string, string> listRow, Base.Web.NDAWebClient client)
        {
            string userAgent = this.RunPage.CurrentUserAgents.GetOneUserAgent();
            client.Headers["User-Agent"] = userAgent; 
            client.Headers.Add("content-type", "application/x-www-form-urlencoded");
        }
        public override void CheckRequestCompleteFile(string webPageText, Dictionary<string, string> listRow)
        {
            if (webPageText.Contains("小区大全") && !webPageText.Contains("访问验证") && webPageText.Trim().EndsWith("</html>"))
            {

            }
            else
            {
                throw new Exception("未完全加载文件.");
            }
        }         

        public override bool AfterAllGrab(IListSheet listSheet)
        {
            try
            {
                this.GetXiaoquInfos(listSheet);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return true;
        }

        private ExcelWriter GetXiaoquExcelWriter()
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("cityCode", 0);
            resultColumnDic.Add("cityName", 1);
            resultColumnDic.Add("level1AreaCode", 2);
            resultColumnDic.Add("level1AreaName", 3);
            resultColumnDic.Add("level2AreaCode", 4);
            resultColumnDic.Add("level2AreaName", 5);
            resultColumnDic.Add("name", 6);
            resultColumnDic.Add("lat", 7);
            resultColumnDic.Add("lng", 8);
            resultColumnDic.Add("address", 9);
            resultColumnDic.Add("sale_num", 10);
            resultColumnDic.Add("build_year", 11);
            resultColumnDic.Add("mid_price", 12);
            resultColumnDic.Add("url", 13);
            string resultFilePath = Path.Combine(exportDir, "安居客小区.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }

        private void GetXiaoquInfos(IListSheet listSheet)
        {
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            ExcelWriter resultEW = this.GetXiaoquExcelWriter();

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"];
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                    try
                    {
                        HtmlNode nameNode = htmlDoc.DocumentNode.SelectSingleNode("//div[@class=\"comm-tit\"]/h1");
                        string name = nameNode.InnerText.Trim();

                        HtmlNode mapLinkNode = htmlDoc.DocumentNode.SelectSingleNode("//div[@class=\"comm-tit\"]/div[@class=\"comm-ad\"]/p/a[@class=\"ad-ico\"]");
                        string lat = "";
                        string lng = "";
                        if (mapLinkNode != null)
                        {
                            string mapLinkUrl = mapLinkNode.GetAttributeValue("href", "");
                            if (mapLinkUrl.Length > 0)
                            {
                                string[] mapLinkParts = mapLinkUrl.Split(new string[] { "&" }, StringSplitOptions.RemoveEmptyEntries);
                                foreach (string mapLinkPart in mapLinkParts)
                                {
                                    if (mapLinkPart.StartsWith("lng="))
                                    {
                                        int startLngIndex = mapLinkPart.IndexOf("=");
                                        lng = mapLinkPart.Substring(startLngIndex + 1);
                                    }
                                    else if (mapLinkPart.StartsWith("lat="))
                                    {
                                        int startLatIndex = mapLinkPart.IndexOf("=");
                                        lat = mapLinkPart.Substring(startLatIndex + 1);
                                    }
                                }
                            }
                        }

                        HtmlNode priceNode = htmlDoc.DocumentNode.SelectSingleNode("//a[@class=\"item main-item\"]/div[@class=\"txt-c\"]/p[@class=\"price\"]");
                        string price = priceNode == null ? "" : priceNode.InnerText.Trim();

                        Dictionary<string, string> f2vs = new Dictionary<string, string>();
                        f2vs.Add("cityCode", row["cityCode"]);
                        f2vs.Add("cityName", row["cityName"]);
                        f2vs.Add("level1AreaCode", row["level1AreaCode"]);
                        f2vs.Add("level1AreaName", row["level1AreaName"]);
                        f2vs.Add("level2AreaCode", row["level2AreaCode"]);
                        f2vs.Add("level2AreaName", row["level2AreaName"]);
                        f2vs.Add("name", name);
                        f2vs.Add("lat", lat);
                        f2vs.Add("lng", lng);
                        f2vs.Add("address", row["address"]);
                        f2vs.Add("sale_num", row["sale_num"]);
                        f2vs.Add("build_year", row["build_year"]);
                        f2vs.Add("mid_price", row["mid_price"]);
                        f2vs.Add("url", detailUrl);
                        resultEW.AddRow(f2vs);
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
            }
            resultEW.SaveToDisk();
        }
    }
}