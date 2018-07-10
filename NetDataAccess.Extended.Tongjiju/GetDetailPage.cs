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
using NetDataAccess.Base.UserAgent;
using System.Net;

namespace NetDataAccess.Extended.Dzdp
{
    public class GetDetailPage : ExternalRunWebPage
    {
        public override void WebRequestHtml_BeforeSendRequest(string pageUrl, Dictionary<string, string> listRow, Base.Web.NDAWebClient client)
        { 
            string userAgent = this.RunPage.CurrentUserAgents.GetOneUserAgent();
            client.Headers["User-Agent"] = userAgent;
            //this.RunPage.InvokeAppendLogText("使用了UserAgent: " + userAgent, LogLevelType.System, true);
        }

        public override bool AfterGrabOneCatchException(string pageUrl, System.Collections.Generic.Dictionary<string, string> listRow, System.Exception ex)
        {
            if (ex.InnerException is WebException)
            {
                WebException webEx = (WebException)ex.InnerException;
                if (webEx.Status == WebExceptionStatus.ProtocolError)
                {
                    HttpWebResponse webRes = (HttpWebResponse)webEx.Response;
                    if (webRes.StatusCode == HttpStatusCode.NotFound)
                    {
                        this.RunPage.InvokeAppendLogText("服务器端不存在此网页(404), pageUrl = " + pageUrl, LogLevelType.Error, true);
                        return true;
                    }
                }
            }
            return false;
        }

        public override void CheckRequestCompleteFile(string webPageText, Dictionary<string, string> listRow)
        {
            string shopId =  listRow["detailPageName"];
            if (webPageText.Contains("shopId: " + shopId) && webPageText.Trim().EndsWith("</html>"))
            {

                //是否包含经纬度
                int latNameBeginIndex = webPageText.IndexOf("shopGlat:");
                string latStr = "";
                string lngStr = "";
                if (latNameBeginIndex > 0)
                {
                    int latBeginIndex = webPageText.IndexOf("\"", latNameBeginIndex);
                    int latEndIndex = webPageText.IndexOf("\"", latBeginIndex + 1);
                    latStr = webPageText.Substring(latBeginIndex + 1, latEndIndex - latBeginIndex - 1);
                }
                int lngNameBeginIndex = webPageText.IndexOf("shopGlng:");
                if (lngNameBeginIndex > 0)
                {
                    int lngBeginIndex = webPageText.IndexOf("\"", lngNameBeginIndex);
                    int lngEndIndex = webPageText.IndexOf("\"", lngBeginIndex + 1);
                    lngStr = webPageText.Substring(lngBeginIndex + 1, lngEndIndex - lngBeginIndex - 1);
                }
                if ((latStr.Length == 0 && lngStr.Length != 0) || (latStr.Length != 0 && lngStr.Length == 0))
                {
                    throw new Exception("经纬度缺失一个，未完全加载文件.");
                }

                //是否包含“口味、环境、服务”

                if (webPageText.Contains("/g134\" itemprop=\"url\"> 茶馆 </a>")
                    ||webPageText.Contains("/g3064\" itemprop=\"url\"> 快照摄影 </a>"))
                {
                    //茶馆
                    //可以没有“口味、环境、服务”

                }
                else
                {
                    if (!webPageText.Contains("comment_score"))
                    {
                        throw new Exception("不包含评论得分，未完全加载文件.");
                    }
                }
            }
            else
            {
                throw new Exception("不包含shopId，或者html不完整，未完全加载文件.");
            }
        }

        public override bool AfterAllGrab(IListSheet listSheet)
        { 
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("city", 0);
            resultColumnDic.Add("distrctName", 1);
            resultColumnDic.Add("shopName", 2);
            resultColumnDic.Add("shopCode", 3);
            resultColumnDic.Add("address", 4);
            resultColumnDic.Add("tel", 5);
            resultColumnDic.Add("shopType", 6);
            resultColumnDic.Add("commentNum", 7);
            resultColumnDic.Add("lat", 8);
            resultColumnDic.Add("lng", 9);
            resultColumnDic.Add("人均", 10);
            resultColumnDic.Add("口味", 11);
            resultColumnDic.Add("环境", 12);
            resultColumnDic.Add("服务", 13);
            string resultFilePath = Path.Combine(exportDir, "大众点评店铺信息.xlsx");
            Dictionary<string, string> resultColumnFormat = new Dictionary<string, string>();
            resultColumnFormat.Add("reviewNum", "#,##0");
            resultColumnFormat.Add("lat", "#,##0.000000");
            resultColumnFormat.Add("lng", "#,##0.000000");
            resultColumnFormat.Add("人均", "#,##0.00");
            resultColumnFormat.Add("环境", "#,##0.0");
            resultColumnFormat.Add("口味", "#,##0.0");
            resultColumnFormat.Add("服务", "#,##0.0");

            CsvWriter resultEW = new CsvWriter(resultFilePath, resultColumnDic);

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;

            Dictionary<string, string> shopDic = new Dictionary<string, string>();

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    try
                    {
                        string url = row[detailPageUrlColumnName];
                        string city = row["city"];
                        string distrctName = row["rName"];
                        string shopName = row["shopName"];
                        string shopCode = row["shopCode"];
                        string shopType = row["gName"];
                        string commentNumStr = row["reviewNum"];
                        Nullable<int> commentNum = commentNumStr == null || commentNumStr.Length == 0 ? (Nullable<int>)null : int.Parse(row["reviewNum"]);
                        Nullable<decimal> lat = null;
                        Nullable<decimal> lng = null;
                        string address = "";
                        string tel = "";
                        Nullable<decimal> renJun = null;
                        Nullable<decimal> kouWei = null;
                        Nullable<decimal> huanJing = null;
                        Nullable<decimal> fuWu = null;



                        HtmlAgilityPack.HtmlDocument pageHtmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                        string pageText = pageHtmlDoc.DocumentNode.InnerHtml;
                         
                        int latNameBeginIndex = pageText.IndexOf("shopGlat:");
                        if (latNameBeginIndex > 0)
                        {
                            int latBeginIndex = pageText.IndexOf("\"", latNameBeginIndex);
                            int latEndIndex = pageText.IndexOf("\"", latBeginIndex + 1);
                            if (latEndIndex - latBeginIndex > 1)
                            {
                                decimal latValue = 0;
                                if (decimal.TryParse(pageText.Substring(latBeginIndex + 1, latEndIndex - latBeginIndex - 1), out latValue))
                                {
                                    lat = latValue;
                                }
                            }
                        }
                        int lngNameBeginIndex = pageText.IndexOf("shopGlng:");
                        if (lngNameBeginIndex > 0)
                        {
                            int lngBeginIndex = pageText.IndexOf("\"", lngNameBeginIndex);
                            int lngEndIndex = pageText.IndexOf("\"", lngBeginIndex + 1);
                            if (lngEndIndex - lngBeginIndex > 1)
                            { 
                                decimal lngValue = 0;
                                if (decimal.TryParse(pageText.Substring(lngBeginIndex + 1, lngEndIndex - lngBeginIndex - 1), out lngValue))
                                {
                                    lng = lngValue;
                                }
                            }
                        }
                        /*
                        HtmlNode preMapScriptNode = pageHtmlDoc.DocumentNode.SelectSingleNode("//div[@class=\"J_midas-4\"]");
                        if (preMapScriptNode != null)
                        {
                            HtmlNode mapScriptNode = preMapScriptNode.PreviousSibling;
                            while (mapScriptNode != null && mapScriptNode.Name != "script")
                            {
                                mapScriptNode = mapScriptNode.PreviousSibling; 
                            }
                            if (mapScriptNode != null)
                            {
                                string scriptString = mapScriptNode.InnerText;
                                int lngBeginIndex = scriptString.LastIndexOf("{lng:") + 5;
                                int lngEndIndex = scriptString.LastIndexOf(",lat:");
                                int latBeginIndex = lngEndIndex + 5;
                                int latEndIndex = scriptString.LastIndexOf("});");
                                lng = decimal.Parse(scriptString.Substring(lngBeginIndex, lngEndIndex - lngBeginIndex));
                                lat = decimal.Parse(scriptString.Substring(latBeginIndex, latEndIndex - latBeginIndex));
                            }
                        }
                         * */

                        HtmlNode addressNode = pageHtmlDoc.DocumentNode.SelectSingleNode("//span[@itemprop=\"street-address\"]");
                        if (addressNode != null)
                        {
                            address = addressNode.Attributes["title"].Value;
                        }

                        HtmlNodeCollection allTelNodes = pageHtmlDoc.DocumentNode.SelectNodes("//span[@itemprop=\"tel\"]");
                        if (allTelNodes != null)
                        {
                            StringBuilder tels = new StringBuilder();
                            foreach (HtmlNode telNode in allTelNodes)
                            {
                                tels.Append((tels.Length == 0 ? "" : ",") + telNode.InnerText);
                            }
                            tel = tels.ToString();
                        }

                        HtmlNodeCollection allBriefNodes = pageHtmlDoc.DocumentNode.SelectNodes("//div[@class=\"brief-info\"]/span");
                        foreach (HtmlNode briefNode in allBriefNodes)
                        {
                            string briefText = briefNode.InnerText;
                            if (briefText.StartsWith("人均:"))
                            {
                                string briefValue = briefText.Substring(3, briefText.Length - 4).Trim();
                                renJun = briefValue.Length == 0 ? (Nullable<decimal>)null : decimal.Parse(briefValue);
                            }
                        }

                        HtmlNodeCollection allScoreNodes = pageHtmlDoc.DocumentNode.SelectNodes("//span[@id=\"comment_score\"]/span");
                        if (allScoreNodes != null)
                        {
                            foreach (HtmlNode scoreNode in allScoreNodes)
                            {
                                string scoreText = scoreNode.InnerText;
                                if (scoreText.StartsWith("口味:"))
                                {
                                    string scoreValue = scoreText.Substring(3).Trim();
                                    kouWei = scoreValue.Length == 0 ? (Nullable<decimal>)null : decimal.Parse(scoreValue);
                                }
                                else if (scoreText.StartsWith("环境:"))
                                {
                                    string scoreValue = scoreText.Substring(3).Trim();
                                    huanJing = scoreValue.Length == 0 ? (Nullable<decimal>)null : decimal.Parse(scoreValue);
                                }
                                else if (scoreText.StartsWith("服务:"))
                                {
                                    string scoreValue = scoreText.Substring(3).Trim();
                                    fuWu = scoreValue.Length == 0 ? (Nullable<decimal>)null : decimal.Parse(scoreValue);
                                }
                            }
                        }

                        Dictionary<string, string> f2vs = new Dictionary<string, string>();
                        f2vs.Add("city", city);
                        f2vs.Add("distrctName", distrctName);
                        f2vs.Add("shopName", shopName);
                        f2vs.Add("shopCode", shopCode);
                        f2vs.Add("address", address);
                        f2vs.Add("shopType", shopType);
                        f2vs.Add("commentNum", commentNum.ToString());
                        f2vs.Add("lat", lat.ToString());
                        f2vs.Add("lng", lng.ToString());
                        f2vs.Add("人均", renJun.ToString());
                        f2vs.Add("tel", tel);
                        f2vs.Add("口味", kouWei.ToString());
                        f2vs.Add("服务", fuWu.ToString());
                        f2vs.Add("环境", huanJing.ToString());
                        resultEW.AddRow(f2vs);
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
            }

            resultEW.SaveToDisk();

            return true;
        } 
    }
}