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

namespace NetDataAccess.Extended.Id.Google
{
    /// <summary>
    /// GetAllThreadPages
    /// </summary>
    [System.Runtime.InteropServices.ComVisibleAttribute(true)]
    public class GetAllThreadPages : ExternalRunWebPage
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

        public override void CheckRequestCompleteFile(string webPageText, Dictionary<string, string> listRow)
        {
            base.CheckRequestCompleteFile(webPageText, listRow);
            TextReader tr = null;

            try
            {
                HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                htmlDoc.LoadHtml(webPageText);

                HtmlNodeCollection loadingNodes = htmlDoc.DocumentNode.SelectNodes("//span[contains(@id, 'message_snippet')]");
                if (loadingNodes != null && loadingNodes.Count != 0)
                {
                    throw new Exception("文档未加载完成");
                }

                HtmlNodeCollection listNodeList = htmlDoc.DocumentNode.SelectNodes("//div[@id=\"tm-tl\"]/div[contains(@class, 'F0XO1GC-nb-W')]");

                if (listNodeList == null || listNodeList.Count == 0)
                {
                    throw new Exception("文档未加载完成");
                }
                else
                {
                    HtmlNode createPostNode = listNodeList[0];
                    string creator = this.GetPostUser(createPostNode);
                    if (creator == null || creator.Length == 0)
                    {
                        throw new Exception("文档未加载完成");
                    }

                    
                    /*
                    HtmlNode lastPostNode = listNodeList[listNodeList.Count - 1];
                    string postTime = this.getPostTime(lastPostNode);
                    if (postTime == null || postTime.Length == 0)
                    {
                        throw new Exception("文档未加载完成");
                    }*/
                }
            }
            catch (Exception ex)
            {
                if (tr != null)
                {
                    tr.Dispose();
                    tr = null;
                }
                throw new Exception("文档未加载完成");
            }
        }

        public override void WebRequestHtml_BeforeSendRequest(string pageUrl, Dictionary<string, string> listRow, Base.Web.NDAWebClient client)
        {
            base.WebRequestHtml_BeforeSendRequest(pageUrl, listRow, client);
            client.Headers.Add("accept-language", "en,zh-CN;q=0.8,zh;q=0.6");
        } 

        public override bool AfterAllGrab(IListSheet listSheet)
        {
            return this.GetItemInfos(listSheet) && this.GetAllItemPosters(listSheet);
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

            List<object[]> threadInfoColumns = new List<object[]>();
            threadInfoColumns.Add(new object[] { "url", null, 70 });
            threadInfoColumns.Add(new object[] { "code", null, 15 });
            threadInfoColumns.Add(new object[] { "title", null, 100 });
            threadInfoColumns.Add(new object[] { "postCount", null, 10 });
            threadInfoColumns.Add(new object[] { "viewCount", null, 10 });
            threadInfoColumns.Add(new object[] { "createTime", null, 20 });
            threadInfoColumns.Add(new object[] { "lastPostTime", null, 20 });
            threadInfoColumns.Add(new object[] { "creator", null, 20 });
            threadInfoColumns.Add(new object[] { "atlassianLink", null, 20 });
            string threadFilePath = Path.Combine(exportDir, "数据3_Id_Google_All.xlsx");
            ExcelWriter theadEw = new ExcelWriter(threadFilePath, "List", threadInfoColumns);
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"];
                string code = row["detailPageName"];
                string title = row["title"];
                string postCount = row["posts"];
                string viewCount = row["views"];

                if (row["giveUpGrab"] != "Y")
                {
                    string localFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir);

                    try
                    {
                        HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                        this.GetItem(htmlDoc, theadEw, detailUrl, code, title, postCount, viewCount);
                    }
                    catch (Exception ex)
                    {
                        this.RunPage.InvokeAppendLogText("读取出错. " + ex.Message + " LocalPath = " + localFilePath, LogLevelType.Error, true); 
                    }
                }
            }
            theadEw.SaveToDisk();
            return succeed;
        }
        /// <summary>
        /// 获取列表页里信息
        /// </summary>
        /// <param name="listSheet"></param>
        /// <returns></returns>
        private bool GetAllItemPosters(IListSheet listSheet)
        {
            bool succeed = true;
            string exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            List<object[]> threadInfoColumns = new List<object[]>();
            threadInfoColumns.Add(new object[] { "url", null, 100 });
            threadInfoColumns.Add(new object[] { "code", null, 20 });
            threadInfoColumns.Add(new object[] { "poster", null, 20 });
            threadInfoColumns.Add(new object[] { "postTime", null, 20 }); 
            string threadFilePath = Path.Combine(exportDir, "数据3_Id_Google_Posters.xlsx");
            ExcelWriter theadEw = new ExcelWriter(threadFilePath, "List", threadInfoColumns);
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
                        HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                        this.GetItemPoster(htmlDoc, theadEw, detailUrl, code );
                    }
                    catch (Exception ex)
                    {
                        this.RunPage.InvokeAppendLogText("读取出错. " + ex.Message + " LocalPath = " + localFilePath, LogLevelType.Error, true);
                        File.Delete(localFilePath);
                    }
                }
            }
            theadEw.SaveToDisk();
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

        private void GetItem(HtmlAgilityPack.HtmlDocument htmlDoc, ExcelWriter threadEw, string detailUrl, string code, string title, string postCount, string viewCount)
        {
            string createTime = "";
            string lastPostTime = "";
            string creator = "";
            string atlassianLink = "";

            HtmlNodeCollection listNodeList = htmlDoc.DocumentNode.SelectNodes("//div[@id=\"tm-tl\"]/div[contains(@class, 'F0XO1GC-nb-W')]");

            if (listNodeList != null && listNodeList.Count > 0)
            {
                HtmlNode createPostNode = listNodeList[0];

                int lastIndex = listNodeList.Count - 1;
                HtmlNode lastPostNode = listNodeList[lastIndex];
                while (this.isDeletedPost(lastPostNode))
                {
                    lastIndex = lastIndex - 1;
                    lastPostNode = listNodeList[lastIndex];
                }

                creator = this.GetPostUser(createPostNode);
                createTime = this.getPostTime(createPostNode);
                lastPostTime = this.getPostTime(lastPostNode);

                List<string> links = new List<string>();
                for (int i = 0; i < listNodeList.Count; i++)
                {
                    this.getAtlassianLink(listNodeList[i], links);
                }
                if (links.Count > 0)
                {
                    atlassianLink = CommonUtil.StringArrayToString(links.ToArray(), "\r\n");
                }

                Dictionary<string, object> itemInfo = new Dictionary<string, object>();
                itemInfo.Add("url", detailUrl);
                itemInfo.Add("code", code);
                itemInfo.Add("title", title);
                itemInfo.Add("postCount", postCount);
                itemInfo.Add("viewCount", viewCount);
                itemInfo.Add("createTime", createTime);
                itemInfo.Add("lastPostTime", lastPostTime);
                itemInfo.Add("creator", creator);
                itemInfo.Add("atlassianLink", atlassianLink);

                threadEw.AddRow(itemInfo);
            }
        }

        private void GetItemPoster(HtmlAgilityPack.HtmlDocument htmlDoc, ExcelWriter threadEw, string detailUrl, string code)
        { 

            HtmlNodeCollection listNodeList = htmlDoc.DocumentNode.SelectNodes("//div[@id=\"tm-tl\"]/div[contains(@class, 'F0XO1GC-nb-W')]");

            if (listNodeList != null && listNodeList.Count > 0)
            {
                foreach (HtmlNode postNode in listNodeList)
                {
                    if (!this.isDeletedPost(postNode))
                    {
                        string poster = this.GetPostUser(postNode);
                        string postTime = this.getPostTime(postNode);

                        Dictionary<string, object> itemInfo = new Dictionary<string, object>();
                        itemInfo.Add("url", detailUrl);
                        itemInfo.Add("code", code);
                        itemInfo.Add("poster", poster);
                        itemInfo.Add("postTime", postTime);

                        threadEw.AddRow(itemInfo);
                    }
                }
            }
        }
        
        private string GetPostUser(HtmlNode postNode)
        {
            try
            {
                string postUser = postNode.SelectSingleNode("./div/div/div/div/div/table/tbody/tr/td[@align='left']/span/span").InnerText.Trim();
                return postUser;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private string getPostTime(HtmlNode postNode)
        {
            try
            {
                string timeStr = postNode.SelectSingleNode("./div/div/div/div/div/table/tbody/tr/td[@align='right']/div[contains(@class, 'F0XO1GC-nb-R')]/span").GetAttributeValue("title", "");
                int fromIndex = timeStr.IndexOf(",");
                int toIndex = timeStr.LastIndexOf(" ");
                timeStr = timeStr.Substring(fromIndex + 1, toIndex - fromIndex).Trim().Replace(",", "").Replace(" at ", " ");
                string formattedTimeStr = this.GetFormatedTimeString(timeStr);
                return formattedTimeStr;
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }
        private bool isDeletedPost(HtmlNode postNode)
        {
            return postNode.InnerText.Contains("message has been deleted.") || postNode.InnerText.Contains("messages have been deleted.");
        }

        private void getAtlassianLink(HtmlNode postNode, List<string> links)
        {
            List<HtmlNode> linkNodes = new List<HtmlNode>();
            this.GetHtmlNodeByAttributeValueContainText(postNode,"href", "idempiere.atlassian", linkNodes);
            if (linkNodes.Count != 0)
            {  
                foreach (HtmlNode linkNode in linkNodes)
                {
                    string link = linkNode.GetAttributeValue("href", "");
                    if (!links.Contains(link))
                    {
                        links.Add(link);
                    }
                } 
            }
        }

    }
}