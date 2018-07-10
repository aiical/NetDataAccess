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
using NetDataAccess.Base.Reader;
using NetDataAccess.Extended.Taobao.Common;
using System.Xml;
using System.Net;
using System.Text.RegularExpressions;

namespace NetDataAccess.Extended.Id.List
{
    /// <summary>
    /// GetAllListPages
    /// </summary>
    [System.Runtime.InteropServices.ComVisibleAttribute(true)]
    public class GetAllActivityPages : ExternalRunWebPage
    {
        public override bool AfterGrabOneCatchException(string pageUrl, Dictionary<string, string> listRow, Exception ex)
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

        public override void WebRequestHtml_BeforeSendRequest(string pageUrl, Dictionary<string, string> listRow, Base.Web.NDAWebClient client)
        {
            base.WebRequestHtml_BeforeSendRequest(pageUrl, listRow, client);
            client.Headers.Add("accept-language", "en,zh-CN;q=0.8,zh;q=0.6");
        } 

        public override bool AfterAllGrab(IListSheet listSheet)
        {
            return this.GetItemInfos(listSheet);
        } 

        /// <summary>
        /// 获取列表页里信息
        /// </summary>
        /// <param name="listSheet"></param>
        /// <returns></returns>
        private bool GetItemInfos(IListSheet listSheet)
        {
            bool succeed = true;
            string exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
             
            List<object[]> activityInfoColumns = new List<object[]>(); 
            activityInfoColumns.Add(new object[] { "code", null, 15 });
            activityInfoColumns.Add(new object[] { "username", null, 20 });
            activityInfoColumns.Add(new object[] { "posttime", null, 20 });
            activityInfoColumns.Add(new object[] { "operate", null, 25 });
            activityInfoColumns.Add(new object[] { "message", null, 100 });
            activityInfoColumns.Add(new object[] { "url", null, 100 });
            string activityFilePath = Path.Combine(exportDir, "数据1_Id_Activity_All.xlsx");
            ExcelWriter activityEw = new ExcelWriter(activityFilePath, "List", activityInfoColumns);
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"];
                string code = row["detailPageName"];

                if (row["giveUpGrab"] != "Y")
                {
                    string localFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir);

                    try
                    {
                        StreamReader tr = new StreamReader(localFilePath, Encoding.UTF8);
                        string webPageHtml = tr.ReadToEnd();
                        string startString = "WRM._unparsedData[\"activity-panel-pipe-id\"]=\"\\\"";
                        string endString = "\\\"\";\nif(window.WRM._dataArrived)window.WRM._dataArrived();";
                        int preStartIndex = webPageHtml.IndexOf(startString) + startString.Length;
                        int startIndex = preStartIndex + startString.Length;
                        int endIndex = webPageHtml.IndexOf(endString, startIndex);
                        if (preStartIndex < 0 || endIndex < 0 || startIndex>=endIndex)
                        {
                        }
                        string allActivityEncodeHtml = webPageHtml.Substring(startIndex, endIndex - startIndex);

                        string allActivityHtml = Regex.Unescape(Regex.Unescape(allActivityEncodeHtml));
                        HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                        htmlDoc.LoadHtml(allActivityHtml);

                        this.GetItem(htmlDoc, activityEw, code, detailUrl);
                    }
                    catch (Exception ex)
                    {
                        this.RunPage.InvokeAppendLogText("读取出错. " + ex.Message + " LocalPath = " + localFilePath, LogLevelType.Error, true);
                    }
                }
            }
            activityEw.SaveToDisk();
            return succeed;
        }

        private String GetFormatedTimeString(String sourceTimeString)
        {
            String destTimeString = "";
            DateTime dt = new DateTime();
            if (DateTime.TryParse(sourceTimeString, out dt))
            {
                destTimeString = dt.ToString("yyyy-MM-dd HH:mm:ss");
            }
            return destTimeString;
        }
         
        private HtmlAgilityPack.HtmlDocument GetHtmlDocByHtml(string html)
        {
            HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
            htmlDoc.LoadHtml(html);
            return htmlDoc;
        }

        private string GetTextByHtmlDoc(HtmlAgilityPack.HtmlDocument htmlDoc)
        {  
            return htmlDoc.DocumentNode.InnerText;
        }
         
        private void GetHtmlNodeByAttributeValueContainText(HtmlNode node, string attributeName, string checkString, List<HtmlNode> checkedNodes)
        { 
            HtmlAttribute attr = node.Attributes[attributeName];
            if (attr != null)
            {
                if (attr.Value.Contains(checkString))
                {
                    checkedNodes.Add(node);
                }
            }

            HtmlNodeCollection childNodes = node.ChildNodes;
            foreach (HtmlNode childNode in childNodes)
            {
                this.GetHtmlNodeByAttributeValueContainText(childNode, attributeName, checkString, checkedNodes);
            }
        }

        private void GetItem(HtmlAgilityPack.HtmlDocument htmlDoc, ExcelWriter activityEw, string code, string pageUrl)
        {
            string username = ""; 
            string posttime = "";
            string operate = "";
            string message = ""; 

            HtmlNodeCollection actNodes = htmlDoc.DocumentNode.SelectNodes("//div[@id=\"issue_actions_container\"]/div");

            if (actNodes != null && actNodes.Count > 0)
            {
                for (int i = 0; i < actNodes.Count; i++)
                {
                    HtmlNode actNode = actNodes[i];
                    HtmlNode actDetailNode = actNode.SelectSingleNode("./div/div[@class=\"action-details\"]");
                    if (actDetailNode == null)
                    {
                        actDetailNode = actNode.SelectSingleNode("./div/div/div[@class=\"action-details\"]");
                    }
                    HtmlNode usernameNode = actDetailNode.SelectSingleNode("./a");
                    if (usernameNode == null)
                    {
                        usernameNode = actDetailNode.SelectSingleNode("./span[@class=\"user-hover user-avatar\"]");
                        username = usernameNode.InnerText.Trim();
                    }
                    else
                    {
                        username = usernameNode.GetAttributeValue("rel", "");
                    }
                    operate = usernameNode.NextSibling.InnerText.Trim().Replace("-","").Trim();

                    HtmlNode postTimeNode = actDetailNode.SelectSingleNode("./span/time[@class=\"livestamp\"]");
                    if (postTimeNode == null)
                    {
                        postTimeNode = actDetailNode.SelectSingleNode("./span/span/time[@class=\"livestamp\"]");
                    }
                    string postTimeStr = "";

                    if (postTimeNode == null)
                    {
                        postTimeStr = actDetailNode.InnerText.Trim();
                    }
                    else
                    {
                        postTimeStr = postTimeNode.GetAttributeValue("datetime", "");
                    }
                    posttime = this.GetFormatedTimeString(postTimeStr);

                    message = actNode.InnerText.Trim();

                    Dictionary<string, object> itemInfo = new Dictionary<string, object>();
                    itemInfo.Add("username", username);
                    itemInfo.Add("code", code);
                    itemInfo.Add("posttime", posttime);
                    itemInfo.Add("operate", operate);
                    itemInfo.Add("message", message.Length > 32765 ? message.Substring(0, 32765) : message);
                    itemInfo.Add("url", pageUrl);

                    activityEw.AddRow(itemInfo);
                }
            }
        }
    }
}