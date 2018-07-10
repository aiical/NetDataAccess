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

namespace NetDataAccess.Extended.Id.Google
{
    /// <summary>
    /// GetAllListPages
    /// </summary>
    [System.Runtime.InteropServices.ComVisibleAttribute(true)]
    public class GetAllDetailPageUrls : ExternalRunWebPage
    {
        public override void WebBrowserHtml_AfterPageLoaded(string pageUrl, Dictionary<string, string> listRow, WebBrowser webBrowser)
        {
            base.WebBrowserHtml_AfterPageLoaded(pageUrl, listRow, webBrowser);
            string tabName = Thread.CurrentThread.ManagedThreadId.ToString();
            this.GetAllItems(webBrowser, tabName); 
        }

        public override bool AfterAllGrab(IListSheet listSheet)
        { 
            string exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"];
                string code = row["detailPageName"];

                if (row["giveUpGrab"] != "Y")
                {
                    string localFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir);

                    StreamReader tr = new StreamReader(localFilePath, Encoding.UTF8);
                    string webPageHtml = tr.ReadToEnd();
                    HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                    htmlDoc.LoadHtml(webPageHtml);
                    HtmlNodeCollection itemNodes = htmlDoc.DocumentNode.SelectNodes("//table[@role=\"list\"]/tbody/tr");
                    this.GetInfos(itemNodes);
                }
            }
            return true;
        }

        /// <summary>
        /// 获取列表页里的店铺信息
        /// </summary>
        /// <param name="listSheet"></param>
        /// <returns></returns>
        private bool GetInfos(HtmlNodeCollection itemNodes)
        {
            bool succeed = true;
            string exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> columnDic = CommonUtil.InitStringIndexDic(new string[]{
                    "detailPageUrl",
                    "detailPageName", 
                    "cookie",
                    "grabStatus", 
                    "giveUpGrab",
                    "title",
                    "posts",
                    "views"});
            string shopFirstPageUrlFilePath = Path.Combine(exportDir, "Id_Google_详情页.xlsx");
            ExcelWriter ew = new ExcelWriter(shopFirstPageUrlFilePath, "List", columnDic, null);

            Dictionary<string, string> nameDic = new Dictionary<string, string>();
            for (int i = 0; i < itemNodes.Count; i++)
            {
                HtmlNode itemNode = itemNodes[i];
                try
                {
                    this.GetItem(itemNode, nameDic, ew);
                }
                catch (Exception ex)
                {
                    this.RunPage.InvokeAppendLogText("获取出错. " + ex.Message, LogLevelType.Error, true);
                }
            }
            ew.SaveToDisk();
            return succeed;
        }

        private void GetItem(HtmlNode itemNode, Dictionary<string, string> nameDic, ExcelWriter ew)
        {
            string detailPageUrl = "";
            string detailPageName = "";
            string title = "";
            string posts = "";
            string views = "";

            HtmlNode linkNode = itemNode.SelectSingleNode("./td/div/div/div/div/div/div/a[contains(@id,\"l_topic_title_\")]");
            string href = linkNode.GetAttributeValue("href", "");
            detailPageName = href.Substring(href.LastIndexOf("/") + 1);
            detailPageUrl = "https://groups.google.com/forum/?hl=en" + href;
            title = CommonUtil.HtmlDecode(linkNode.InnerText.Trim()).Trim();
            HtmlNodeCollection spanNodes = itemNode.SelectNodes("./td/div/div/div/div/div/div/span");
            foreach (HtmlNode spanNode in spanNodes)
            {
                string text = spanNode.InnerText.Trim();
                if (text.EndsWith(" posts") || text.EndsWith(" post"))
                {
                    posts = text.Substring(0, text.IndexOf(" "));
                }
                else if (text.EndsWith(" views") || text.EndsWith(" view"))
                {
                    views = text.Substring(0, text.IndexOf(" "));
                }
            }


            Dictionary<string, object> itemInfo = new Dictionary<string, object>();
            if (!nameDic.ContainsKey(detailPageName))
            {
                nameDic.Add(detailPageName, null);
                itemInfo.Add("detailPageUrl", detailPageUrl);
                itemInfo.Add("detailPageName", detailPageName);
                itemInfo.Add("title", title);
                itemInfo.Add("posts", posts);
                itemInfo.Add("views", views);
                ew.AddRow(itemInfo);
            }
        }

        private void GetAllItems(WebBrowser webBrowser, string tabName)
        {
            Thread.Sleep(3000);
            string scriptMethodCode = "function getMoreItems(){"
                + "var containerDiv = document.getElementById('Header-container').nextElementSibling;"
                + "containerDiv.firstChild.scrollTop = containerDiv.firstChild.scrollHeight + 100000;"
                + "}"
                + "function getItemHeight(){"
                + "var containerDiv = document.getElementById('Header-container').nextElementSibling;"
                + "return containerDiv.firstChild.scrollHeight;"
                + "}";
            this.RunPage.InvokeAddScriptMethod(webBrowser, scriptMethodCode, null);

            int itemHeight = this.GetItemHeight(webBrowser, tabName);
            int retryTime = 0;
            while (itemHeight != lastTimeItemHeight || retryTime < 20)
            {
                if (itemHeight == lastTimeItemHeight)
                {
                    retryTime++;
                }
                else
                {
                    retryTime = 0;
                    lastTimeItemHeight = itemHeight;

                    this.RunPage.InvokeDoScriptMethod(webBrowser, "getMoreItems", null);
                }
                Thread.Sleep(3000);
                itemHeight = this.GetItemHeight(webBrowser, tabName);
            } 
        }

        private int lastTimeItemHeight = 0;
        public int GetItemHeight(WebBrowser webBrowser,string tabName)
        {
            int itemHeight = (int)this.RunPage.InvokeDoScriptMethod(webBrowser, "getItemHeight", null);
            return itemHeight;
        }
    }
}