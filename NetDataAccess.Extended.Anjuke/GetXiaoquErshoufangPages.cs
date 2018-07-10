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
    public class GetXiaoquErshoufangPages : ExternalRunWebPage
    {
        public override void WebRequestHtml_BeforeSendRequest(string pageUrl, Dictionary<string, string> listRow, Base.Web.NDAWebClient client)
        {
            string userAgent = this.RunPage.CurrentUserAgents.GetOneUserAgent();
            client.Headers["User-Agent"] = userAgent;
            client.Headers.Add("x-request-with", "XMLHttpRequest");
            client.Headers.Add("cookie", "ctid=" + DateTime.Now.Millisecond);
        }

        public override bool AfterGrabOneCatchException(string pageUrl, Dictionary<string, string> listRow, Exception ex)
        {
            if (ex is EmptyFileException)
            {
                return true;
            }
            else if (ex is GrabRequestException)
            {
                this.RunPage.InvokeAppendLogText(ex.Message, LogLevelType.Error, true);
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
                else
                {
                    return false;
                }
            } 
            return false;
        }

        public override void CheckRequestCompleteFile(string webPageText, Dictionary<string, string> listRow)
        {
            if (webPageText.Contains("article pro-detail") && webPageText.Trim().Contains("</div>"))
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
               // this.GetXiaoquInfos(listSheet);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return true;
        }

        private ExcelWriter GetExcelWriter(int fileIndex)
        { 
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3); 
            resultColumnDic.Add("xiaoquname", 6);  
            resultColumnDic.Add("xiaoquurl", 13);

            string resultFilePath = Path.Combine(exportDir, "安居客在售二手房_" + fileIndex.ToString() + ".xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }

        private void GetXiaoquInfos(IListSheet listSheet)
        {
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            ExcelWriter resultEW = null;
            int fileIndex = 1;

            Dictionary<string, string> fangLinkUrlDic = new Dictionary<string, string>();

            for (int i = 0; i < listSheet.RowCount; i++) 
            {
                if (resultEW == null || resultEW.RowCount > 500000)
                {
                    if (resultEW != null)
                    {
                        resultEW.SaveToDisk();
                    }
                    resultEW = this.GetExcelWriter(fileIndex);
                    fileIndex++;
                }

                Dictionary<string, string> row = listSheet.GetRow(i); 
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                    try
                    {
                        HtmlNodeCollection fangNodeList = htmlDoc.DocumentNode.SelectNodes("//div[@class=\"baseinfo\"]/a");

                        if (fangNodeList != null)
                        {
                            foreach (HtmlNode fangNode in fangNodeList)
                            {
                                string fangLinkUrl = fangNode.GetAttributeValue("href", "");
                                if (!fangLinkUrlDic.ContainsKey(fangLinkUrl))
                                {
                                    fangLinkUrlDic.Add(fangLinkUrl, null);
                                    Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                    f2vs.Add("detailPageUrl", fangLinkUrl);
                                    f2vs.Add("detailPageName", fangLinkUrl);
                                    f2vs.Add("xiaoquname", row["xiaoquName"]);
                                    f2vs.Add("xiaoquurl", row["xiaoquUrl"]);

                                    resultEW.AddRow(f2vs);
                                }
                            }
                        } 
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