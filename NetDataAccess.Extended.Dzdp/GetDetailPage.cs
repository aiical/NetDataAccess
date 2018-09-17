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
            if (webPageText.Length == 0)
            {
                throw new Exception("返回的文件为空");
            }
            else
            {
                string shopId = listRow["detailPageName"];
                if (webPageText.Contains("shopId: \"" + shopId) && webPageText.Trim().EndsWith("</html>"))
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
                        || webPageText.Contains("/g3064\" itemprop=\"url\"> 快照摄影 </a>"))
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
        }

        public override bool AfterAllGrab(IListSheet listSheet)
        { 
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("city", 0);
            resultColumnDic.Add("gName", 1);
            resultColumnDic.Add("rName", 2);
            resultColumnDic.Add("shopName", 3);
            resultColumnDic.Add("reviewNum", 4);
            resultColumnDic.Add("serviceRating", 5);
            resultColumnDic.Add("environmentRating", 6);
            resultColumnDic.Add("tasteRating", 7);
            resultColumnDic.Add("address", 8);
            resultColumnDic.Add("lat", 9);
            resultColumnDic.Add("lng", 10); 
            string resultFilePath = Path.Combine(exportDir, "大众点评店铺信息.xlsx");
            Dictionary<string, string> resultColumnFormat = new Dictionary<string, string>();
            resultColumnFormat.Add("reviewNum", "#,##0");
            resultColumnFormat.Add("lat", "#,##0.000000");
            resultColumnFormat.Add("lng", "#,##0.000000");
            resultColumnFormat.Add("serviceRating", "#,##0.00");
            resultColumnFormat.Add("environmentRating", "#,##0.0");
            resultColumnFormat.Add("tasteRating", "#,##0.0");

            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic);

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
                        Nullable<decimal> lat = null;
                        Nullable<decimal> lng = null; 



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

                        Dictionary<string, object> f2vs = new Dictionary<string, object>();
                        f2vs.Add("city", row["city"]);
                        f2vs.Add("gName", row["gName"]);
                        f2vs.Add("rName", row["rName"]);
                        f2vs.Add("shopName", row["shopName"]);
                        f2vs.Add("reviewNum", row["reviewNum"]);
                        f2vs.Add("serviceRating", row["serviceRating"]);
                        f2vs.Add("environmentRating", row["environmentRating"]);
                        f2vs.Add("tasteRating", row["tasteRating"]);
                        f2vs.Add("address", row["address"]);
                        f2vs.Add("lat", lat);
                        f2vs.Add("lng", lng); 
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