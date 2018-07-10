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

namespace NetDataAccess.Extended.Id.List
{
    /// <summary>
    /// GetAllListPages
    /// </summary>
    [System.Runtime.InteropServices.ComVisibleAttribute(true)]
    public class GetAllDetailPages : ExternalRunWebPage
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

            List<object[]> columns = new List<object[]>();
            columns.Add(new object[] { "id", null, 6 });
            columns.Add(new object[] { "code", null, 15 });
            columns.Add(new object[] { "name", null, 50 });
            columns.Add(new object[] { "type", null, 15  });
            columns.Add(new object[] { "status", null, 15 });
            columns.Add(new object[] { "priority", null, 12 });
            columns.Add(new object[] { "resolution", null, 15 });
            columns.Add(new object[] { "affects Version/s", null, 8 });
            columns.Add(new object[] { "fix Version/s", null, 8 });
            columns.Add(new object[] { "component/s", null, 8 });
            columns.Add(new object[] { "labels", null, 8 });
            columns.Add(new object[] { "assignee", null, 20 });
            columns.Add(new object[] { "reporter", null, 20 });
            columns.Add(new object[] { "testby", null, 20 });
            columns.Add(new object[] { "votes", null, 8 });
            columns.Add(new object[] { "watchers", null, 20 });
            columns.Add(new object[] { "created", null, 20});
            columns.Add(new object[] { "updated", null, 20 });
            columns.Add(new object[] { "resolved", null, 8 });
            columns.Add(new object[] { "environment", null, 8 });
            columns.Add(new object[] { "description", null, 50 });
            columns.Add(new object[] { "gotoGoogleInDescription", null, 40 });
            columns.Add(new object[] { "comment", null, 40 });
            columns.Add(new object[] { "gotoGoogleInComment", null, 40 });
            columns.Add(new object[] { "url", null, 100 });
            string detailFilePath = Path.Combine(exportDir, "数据1_Id详情.xlsx");
            ExcelWriter ew = new ExcelWriter(detailFilePath, "List", columns);


            Dictionary<string, int> activityColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                    "detailPageUrl",
                    "detailPageName", 
                    "cookie",
                    "grabStatus", 
                    "giveUpGrab" }); 
            string activityFilePath = Path.Combine(exportDir, "Id_Activity详情.xlsx");
            ExcelWriter activityEw = new ExcelWriter(activityFilePath, "List", activityColumnDic);
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"];

                if (row["giveUpGrab"] != "Y")
                {
                    string localFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir);

                    try
                    {
                        XmlDocument xmlDoc = new XmlDocument();
                        xmlDoc.Load(localFilePath);
                        this.GetItem(xmlDoc, ew, activityEw);
                    }
                    catch (Exception ex)
                    {
                        this.RunPage.InvokeAppendLogText("读取出错. " + ex.Message + " LocalPath = " + localFilePath, LogLevelType.Error, true);
                    }
                }
            }
            ew.SaveToDisk();
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

        private String GetChildNodeInnerText(XmlNode parentNode, string path)
        {
            XmlNode node = parentNode.SelectSingleNode(path);
            return node == null ? "" : node.InnerText.Trim();
        }

        private String GetChildNodeAttributeValue(XmlNode parentNode, string path, string attributeName)
        {
            XmlNode node = parentNode.SelectSingleNode(path);
            string value = "";
            if (node != null)
            {
                XmlAttribute attri = node.Attributes[attributeName];
                value = attri == null ? "" : attri.Value;
            }
            return value; 
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

        private string GetCustomFieldValue(XmlDocument xmlDoc, String fieldName)
        {
            StringBuilder valueStrings = new StringBuilder();
            XmlNodeList customFieldNodes = xmlDoc.DocumentElement.SelectNodes("./channel/item/customfields/customfield");
            if (customFieldNodes != null)
            {
                foreach (XmlNode customFieldNode in customFieldNodes)
                {
                    if (customFieldNode.SelectSingleNode("./customfieldname").InnerText.Trim() == "Tested By")
                    {
                        XmlNodeList valueNodes = customFieldNode.SelectNodes("./customfieldvalues/customfieldvalue");
                        for (int i = 0; i < valueNodes.Count; i++)
                        {
                            XmlNode valueNode = valueNodes[i];
                            XmlCDataSection cdataNode = (XmlCDataSection)valueNode.ChildNodes[0];

                            if (i > 0)
                            {
                                valueStrings.Append(";");
                            }
                            valueStrings.Append(cdataNode.InnerText.Trim());
                        }
                    }
                }
            }
            return valueStrings.ToString();
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

        private void GetItem(XmlDocument xmlDoc, ExcelWriter ew, ExcelWriter activityEw)
        {
            string id = "";
            string code = "";
            string name = "";
            string type = "";
            string status = "";
            string priority = "";
            string resolution = "";
            string affectsVersion = "";
            string fixVersion = "";
            string component = "";
            string labels = "";
            string assignee = "";
            string reporter = "";
            string testby = "";
            string votes = "";
            string watchers = "";
            string created = "";
            string updated = "";
            string resolved = "";
            string url = "";
            string environment = "";
            string description = "";
            string gotoGoogleInDescription = "";
            string comment = "";
            string gotoGoogleInComment = "";
            XmlNode itemNode = xmlDoc.DocumentElement.SelectSingleNode("./channel/item");

            id = itemNode.SelectSingleNode("./key").Attributes["id"].Value;
            code = this.GetChildNodeInnerText(itemNode, "./key");
            name = this.GetChildNodeInnerText(itemNode, "./summary");
            type = this.GetChildNodeInnerText(itemNode, "./type");
            status = this.GetChildNodeInnerText(itemNode, "./status");
            priority = this.GetChildNodeInnerText(itemNode, "./priority");
            resolution = this.GetChildNodeInnerText(itemNode, "./resolution");
            affectsVersion = this.GetChildNodeInnerText(itemNode, "./version");
            fixVersion = this.GetChildNodeInnerText(itemNode, "./fixVersion");
            XmlNodeList componentNodes = itemNode.SelectNodes("./component");
            if (componentNodes != null && componentNodes.Count > 0)
            {
                StringBuilder componentStr = new StringBuilder();
                for (int i = 0; i < componentNodes.Count; i++)
                {
                    if (i > 0)
                    {
                        componentStr.Append(";");
                    }

                    XmlNode componentNode = componentNodes[i];
                    String com = componentNode.InnerText.Trim();
                    componentStr.Append(com);
                }
                component = componentStr.ToString();
            }

            XmlNodeList labelNodes = itemNode.SelectNodes("./labels");
            if (labelNodes != null && labelNodes.Count > 0)
            {
                StringBuilder labelStr = new StringBuilder();
                for (int i = 0; i < labelNodes.Count; i++)
                {
                    if (i > 0)
                    {
                        labelStr.Append(";");
                    }

                    XmlNode labelNode = labelNodes[i];
                    String label = labelNode.InnerText.Trim();
                    labelStr.Append(label);
                }
                labels = labelStr.ToString();
            }
            string assigneeValue =  this.GetChildNodeAttributeValue(itemNode, "./assignee", "username");
            assignee = assigneeValue == "-1"? "" : assigneeValue;
            string reporterValue = this.GetChildNodeAttributeValue(itemNode, "./reporter", "username");
            reporter = reporterValue == "-1" ? "" : reporterValue;
            votes = this.GetChildNodeInnerText(itemNode, "./votes");
            watchers = this.GetChildNodeInnerText(itemNode, "./watches");
            created = this.GetFormatedTimeString(this.GetChildNodeInnerText(itemNode, "./created"));
            updated = this.GetFormatedTimeString(this.GetChildNodeInnerText(itemNode, "./updated"));
            resolved = this.GetFormatedTimeString(this.GetChildNodeInnerText(itemNode, "./resolved"));
            environment = this.GetChildNodeInnerText(itemNode, "./environment");
            testby = this.GetCustomFieldValue(xmlDoc, "Test by");
            url = "https://idempiere.atlassian.net/browse/" + code;

            string descriptionHtml = CommonUtil.HtmlDecode(this.GetChildNodeInnerText(itemNode, "./description"));
            HtmlAgilityPack.HtmlDocument descriptionHtmlDoc = this.GetHtmlDocByHtml(descriptionHtml);
            description = this.GetTextByHtmlDoc(descriptionHtmlDoc);
            List<HtmlNode> gotoGoogleNodesInDescription = new List<HtmlNode>();
            this.GetHtmlNodeByAttributeValueContainText(descriptionHtmlDoc.DocumentNode, "href", "groups.google.com/forum", gotoGoogleNodesInDescription);
            if (gotoGoogleNodesInDescription != null && gotoGoogleNodesInDescription.Count > 0)
            {
                List<string> gotoGoogleInDescriptionList = new List<string>();
                for (int i = 0; i < gotoGoogleNodesInDescription.Count; i++)
                { 
                    HtmlNode gotoGoogleNodeInDescription = gotoGoogleNodesInDescription[i];

                    String href = gotoGoogleNodeInDescription.Attributes["href"].Value;
                    if (!gotoGoogleInDescriptionList.Contains(href))
                    {
                        gotoGoogleInDescriptionList.Add(href);
                    }
                }
                gotoGoogleInDescription = CommonUtil.StringArrayToString(gotoGoogleInDescriptionList.ToArray(), "\r\n");
            }


            XmlNodeList commentNodes = itemNode.SelectNodes("./comments/comment");
            if (commentNodes != null && commentNodes.Count > 0)
            {
                StringBuilder commentStr = new StringBuilder();
                List<HtmlNode> gotoGoogleNodesInComment = new List<HtmlNode>();
                for (int i = 0; i < commentNodes.Count; i++)
                {
                    if (i > 0)
                    {
                        commentStr.Append("\r\n");
                    }

                    XmlNode commentNode = commentNodes[i];

                    String author = commentNode.Attributes["author"].Value;

                    String postTime = this.GetFormatedTimeString(commentNode.Attributes["created"].Value);

                    HtmlAgilityPack.HtmlDocument commentHtmlDoc = this.GetHtmlDocByHtml(CommonUtil.HtmlDecode(commentNode.InnerText));
                    string c = this.GetTextByHtmlDoc(commentHtmlDoc);
                    commentStr.Append("(" + (i + 1).ToString() + ") " + author + "," + postTime + ":\r\n" + c);


                    this.GetHtmlNodeByAttributeValueContainText(descriptionHtmlDoc.DocumentNode, "href", "groups.google.com/forum", gotoGoogleNodesInComment);

                }
                comment = commentStr.ToString();


                if (gotoGoogleNodesInComment != null && gotoGoogleNodesInComment.Count > 0)
                {
                    List<string> gotoGoogleInCommentList = new List<string>();
                    for (int i = 0; i < gotoGoogleNodesInComment.Count; i++)
                    {
                        HtmlNode gotoGoogleNodeInComment = gotoGoogleNodesInComment[i];

                        String href = gotoGoogleNodeInComment.Attributes["href"].Value;
                        if (!gotoGoogleInCommentList.Contains(href))
                        {
                            gotoGoogleInCommentList.Add(href);
                        }
                    }
                    gotoGoogleInComment = CommonUtil.StringArrayToString(gotoGoogleInCommentList.ToArray(), "\r\n");
                }

            }

            Dictionary<string, object> itemInfo = new Dictionary<string, object>();
            itemInfo.Add("id", id);
            itemInfo.Add("code", code);
            itemInfo.Add("name", name);
            itemInfo.Add("type", type);
            itemInfo.Add("status", status);
            itemInfo.Add("priority", priority);
            itemInfo.Add("resolution", resolution);
            itemInfo.Add("affects Version/s", affectsVersion);
            itemInfo.Add("fix Version/s", fixVersion);
            itemInfo.Add("component/s", component);
            itemInfo.Add("labels", labels);
            itemInfo.Add("assignee", assignee);
            itemInfo.Add("reporter", reporter);
            itemInfo.Add("testby", testby);
            itemInfo.Add("votes", votes);
            itemInfo.Add("watchers", watchers);
            itemInfo.Add("created", created);
            itemInfo.Add("updated", updated);
            itemInfo.Add("description", description);
            itemInfo.Add("gotoGoogleInDescription", gotoGoogleInDescription);
            itemInfo.Add("comment", comment);
            itemInfo.Add("gotoGoogleInComment", gotoGoogleInComment);
            itemInfo.Add("url", url);

            ew.AddRow(itemInfo);


            string activityUrl = "https://idempiere.atlassian.net/browse/" + code + "?page=com.atlassian.jira.plugin.system.issuetabpanels:all-tabpanel&_=1507468385511";
            Dictionary<string, object> activityItemInfo = new Dictionary<string, object>();
            activityItemInfo.Add("detailPageUrl", activityUrl);
            activityItemInfo.Add("detailPageName", code);
            activityEw.AddRow(activityItemInfo);
        }
    }
}