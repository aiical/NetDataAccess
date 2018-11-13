using HtmlAgilityPack;
using NetDataAccess.Base.Browser;
using NetDataAccess.Base.Common;
using NetDataAccess.Base.Config;
using NetDataAccess.Base.DB;
using NetDataAccess.Base.EnumTypes;
using NetDataAccess.Base.Reader;
using NetDataAccess.Base.UI;
using NetDataAccess.Base.Web;
using NetDataAccess.Base.Writer;
using NetDataAccess.Extended.ShangBiao.Common;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace NetDataAccess.Extended.ShangBiao
{
    [System.Runtime.InteropServices.ComVisibleAttribute(true)]
    public class GetInfoPageByCompanyKeyword : ExternalRunWebPage
    {
        public override void GetDataByOtherAccessType(Dictionary<string, string> listRow)
        {
            try
            {
                string pageMarkUrl = listRow[SysConfig.DetailPageUrlFieldName];
                string pageName = listRow[SysConfig.DetailPageNameFieldName];
                string pageUrl = listRow["pageUrl"];
                string keyword = listRow["keyword"];

                IWebBrowser webBrowser = this.OpenInitPage(pageUrl);
                IWebBrowser popBrowser = null;//this.OpenPopPage();

                CefSharp.ICookieManager cookieManager = CefSharp.Cef.GetGlobalCookieManager(); 
                
                this.GoToSearchPage(webBrowser);

                this.DoSearch(webBrowser, keyword, popBrowser);
                 
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        private IWebBrowser OpenInitPage(string pageUrl)
        {
            try
            {
                string tabName = "ShangBiaoPage";
                IWebBrowser webBrowser = this.RunPage.ShowWebPage(pageUrl, tabName, 30000, false, WebBrowserType.Chromium, true);
                this.RunPage.ShowTabPage("ShangBiaoPage");

                if (!webBrowser.Loaded())
                {
                    throw new Exception("页面加载失败");
                }
                return webBrowser;
            }
            catch (Exception ex)
            {
                throw new Exception("页面处理失败. pageUrl = " + pageUrl, ex);
            }
        }


        private IWebBrowser OpenPopPage()
        {
            try
            {
                string tabName = "PopPage";
                IWebBrowser webBrowser = this.RunPage.ShowWebPage("http://wsjs.saic.gov.cn", tabName, 30000, false, WebBrowserType.Chromium, true);
                this.RunPage.ShowTabPage("PopPage");

                if (!webBrowser.Loaded())
                {
                    throw new Exception("页面加载失败");
                }
                return webBrowser;
            }
            catch (Exception ex)
            {
                throw new Exception("加载弹出页面出错", ex);
            }
        }

        private void GoToSearchPage(IWebBrowser webBrowser)
        {
            int interval = 2000;
            int waitingTimeout = 30000;
            int totalWaiting = 0;
            string html = this.RunPage.InvokeGetPageHtml(webBrowser);
            while ((html == null || !html.Contains("商标综合查询")) && waitingTimeout >= totalWaiting)
            {
                totalWaiting += interval;
                Thread.Sleep(interval);
                html = this.RunPage.InvokeGetPageHtml(webBrowser);
            }
            if (totalWaiting > waitingTimeout)
            {
                throw new Exception("跳转到‘商标综合查询’按钮页面超时");
            }
            Thread.Sleep(interval);

            string scriptCode = "var pNodes = document.getElementsByTagName('p');"
                + "for(var i = 0; i < pNodes.length; i++){"
                + "var pNode = pNodes[i];"
                + "if(pNode.innerText == '商标综合查询'){"
                + "pNode.click();"
                + "break;"
                + "}"
                + "}";
            this.RunPage.InvokeAddScriptMethod(webBrowser, scriptCode);
            totalWaiting = 0;
            while ((html == null || !html.Contains("申请人名称（中文）")) && waitingTimeout >= totalWaiting)
            {
                totalWaiting += interval;
                Thread.Sleep(interval);
                html = this.RunPage.InvokeGetPageHtml(webBrowser);
            }
            if (totalWaiting > waitingTimeout)
            {
                throw new Exception("跳转到搜索页面超时");
            }

            //等待页面渲染
            Thread.Sleep(5000);
        }

        private void DoSearch(IWebBrowser webBrowser, string keyword, IWebBrowser popBrowser)
        {
            int interval = 2000;
            int waitingTimeout = 30000;
            int totalWaiting = 0;
            for (int i = 0; i < keyword.Length; i++)
            {
                string partKeyword = keyword[i].ToString();
                string scriptCode = "var inputSqrNodes = document.getElementsByName('request:hnc');"
                    + "inputSqrNodes[0].focus();"
                    + "inputSqrNodes[0].value = inputSqrNodes[0].value + '" + partKeyword + "';";
                this.RunPage.InvokeAddScriptMethod(webBrowser, scriptCode);
                string html = "";

                this.RunPage.InvokeScrollDocumentMethod(webBrowser, new Point(100, 400));

                Thread.Sleep(interval);
            }
            //LifeSpanHandler lifeSpanHandler = new LifeSpanHandler();
            //lifeSpanHandler.PopBrowser = (ChromiumRunWebBrowser)popBrowser;
            //((ChromiumRunWebBrowser)webBrowser).LifeSpanHandler = lifeSpanHandler;

            string scriptSubmitCode = "document.getElementById('_searchButton').focus();document.getElementById('_searchButton').click()";
            this.RunPage.InvokeAddScriptMethod(webBrowser, scriptSubmitCode);

            this.RunPage.ShowTabPage("PopPage");

            //html = this.RunPage.InvokeGetPageHtml(popBrowser); 
              waitingTimeout = 5000; 
            while (waitingTimeout >= totalWaiting)
            {
                totalWaiting += interval;
                Thread.Sleep(interval);
            }
            if (totalWaiting > waitingTimeout)
            {
                //throw new Exception("跳转到搜索结果页面超时");
            }
        }

        void GetInfoPageByCompanyKeyword_FrameLoadStart(object sender, CefSharp.FrameLoadStartEventArgs e)
        {
            throw new NotImplementedException();
        }
        
        private void webBrowserSearchCompleted(IWebBrowser webBrowser)
        { 
            //if (webBrowser.ReadyState == WebBrowserReadyState.Complete && !webBrowser.IsBusy)
            if(webBrowser.Loaded())
            { 
            }
        } 
        private List<string> GetListPageItems(string keywords, IWebBrowser webBrowser)
        {
            List<string> itemList = new List<string>();
            string html = this.RunPage.InvokeGetPageHtml(webBrowser);
            HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
            htmlDoc.LoadHtml(html);
            HtmlNodeCollection itemNodes = htmlDoc.DocumentNode.SelectNodes("//div[@class=\"result-list-record\"]");
            for (int i = 0; i < itemNodes.Count; i++)
            {
                HtmlNode itemNode = itemNodes[i];
                string itemName = this.GetListPageItem(keywords, itemNode);
                itemList.Add(itemName);
            }
            return itemList;
        }

        private string GetListPageItem(string keywords, HtmlNode itemNode)
        {
            HtmlNode itemTitleNode = itemNode.SelectSingleNode("./h3/a");
            if (itemTitleNode == null)
            {
                itemTitleNode = itemNode.SelectSingleNode("./div[@class=\"display-info\"]/h3/a");
            }
            string itemTitle = CommonUtil.HtmlDecode(itemTitleNode.InnerText).Trim();
            HtmlNodeCollection displayInfoNodes = itemNode.SelectNodes("./div[@class=\"display-info\"]/span[@class=\"standard-view-style\"]");
            List<string> displayInfos = new List<string>();
            if (displayInfoNodes != null)
            {
                for (int i = 0; i < displayInfoNodes.Count; i++)
                {
                    HtmlNode displayInfoNode = displayInfoNodes[i];
                    displayInfos.Add(CommonUtil.HtmlDecode(displayInfoNode.InnerText).Trim());
                }
            }

            string itemName = itemTitle + ", " + CommonUtil.StringArrayToString(displayInfos.ToArray(), ", ");

            string filePath = this.GetFilePath(keywords, itemName);
            if (!File.Exists(filePath))
            {
                string baseInfoPageUrl = itemTitleNode.GetAttributeValue("href", "");

                HtmlNode htmlLinkNode = itemNode.SelectSingleNode("./div[@class=\"display-info\"]/div[@class=\"record-formats-wrapper externalLinks\"]/span[@class=\"record-formats\"]/a[@class=\"record-type html-ft\"]");
                HtmlNode pdfLinkNode = itemNode.SelectSingleNode("./div[@class=\"display-info\"]/div[@class=\"record-formats-wrapper externalLinks\"]/span[@class=\"record-formats\"]/a[@class=\"record-type pdf-ft\"]");
                string htmlPageUrl = htmlLinkNode == null ? "" : htmlLinkNode.GetAttributeValue("href", "");
                string pdfPageUrl = pdfLinkNode == null ? "" : pdfLinkNode.GetAttributeValue("href", "");

                if (baseInfoPageUrl.Length > 0)
                {
                    baseInfoPageUrl = CommonUtil.UrlDecodeSymbolAnd(baseInfoPageUrl);
                    this.DownloadBaseInfoPage(baseInfoPageUrl, keywords, itemName);
                }
                if (htmlPageUrl.Length > 0)
                {
                    htmlPageUrl = CommonUtil.UrlDecodeSymbolAnd(htmlPageUrl);
                    this.DownloadHtmlPage(htmlPageUrl, keywords, itemName);
                }
                if (pdfPageUrl.Length > 0)
                {
                    pdfPageUrl = CommonUtil.UrlDecodeSymbolAnd(pdfPageUrl);
                    this.DownloadPdfPage(pdfPageUrl, keywords, itemName);
                }

                this.RunPage.SaveFile(itemName, filePath, Encoding.UTF8);
            }
            return itemName;
        }

        private void DownloadBaseInfoPage(string pageUrl, string keywords, string itemName)
        {
            string filePath = this.GetFilePath(keywords, itemName) + "_baseInfo";
            if (!File.Exists(filePath))
            {
                IWebBrowser webBrowser = this.RunPage.ShowWebPage(pageUrl, "baseInfo", 30000, false, WebBrowserType.Chromium, true);

                int interval = 4000;
                int waitingTimeout = 30000;
                int totalWaiting = 0;
                string html = "";
                while ((html == null || !html.Contains("详细记录") ) && waitingTimeout >= totalWaiting)
                {
                    totalWaiting += interval;
                    Thread.Sleep(interval);
                    html = this.RunPage.InvokeGetPageHtml(webBrowser);
                }
                if (totalWaiting > waitingTimeout)
                {
                    throw new Exception("页面加载失败_baseInfo, keywords = " + keywords + ", itemName = " + itemName + ", pageUrl = " + pageUrl);
                }
                this.RunPage.SaveFile(html, filePath, Encoding.UTF8);
            }
        }
        private void DownloadHtmlPage(string pageUrl, string keywords, string itemName)
        {
            string filePath = this.GetFilePath(keywords, itemName) + "_html";
            if (!File.Exists(filePath))
            {
                IWebBrowser webBrowser = this.RunPage.ShowWebPage(pageUrl, "html", 30000, false, WebBrowserType.Chromium, true);

                int interval = 4000;
                int waitingTimeout = 30000;
                int totalWaiting = 0;
                string html = "";
                while ((html == null || !html.Contains("HTML 全文")) && waitingTimeout >= totalWaiting)
                {
                    totalWaiting += interval;
                    Thread.Sleep(interval);
                    html = this.RunPage.InvokeGetPageHtml(webBrowser);
                }
                if (totalWaiting > waitingTimeout)
                {
                    throw new Exception("页面加载失败_html, keywords = " + keywords + ", itemName = " + itemName + ", pageUrl = " + pageUrl);
                }
                this.RunPage.SaveFile(html, filePath, Encoding.UTF8);
            }
        }
        private void DownloadPdfPage(string pageUrl, string keywords, string itemName)
        {
            try
            {
                string filePath = this.GetFilePath(keywords, itemName) + "_pdf";
                if (!File.Exists(filePath))
                {
                    IWebBrowser webBrowser = this.RunPage.ShowWebPage(pageUrl, "pdf", 30000, false, WebBrowserType.Chromium, true);

                    int interval = 4000;
                    int waitingTimeout = 30000;
                    int totalWaiting = 0;
                    string html = "";
                    while ((html == null || !html.Contains("id=\"pdfIframe\"")) && waitingTimeout >= totalWaiting)
                    {
                        totalWaiting += interval;
                        Thread.Sleep(interval);
                        html = this.RunPage.InvokeGetPageHtml(webBrowser);
                    }
                    if (totalWaiting > waitingTimeout)
                    {
                        throw new Exception("页面加载失败_pdf, keywords = " + keywords + ", itemName = " + itemName + ", pageUrl = " + pageUrl);
                    }

                    totalWaiting = 0;
                    HtmlAgilityPack.HtmlDocument contentHtmlDoc = new HtmlAgilityPack.HtmlDocument();
                    contentHtmlDoc.LoadHtml(html);
                    string contentPageUrl = CommonUtil.UrlDecodeSymbolAnd(contentHtmlDoc.DocumentNode.SelectSingleNode("//iframe[@id=\"pdfIframe\"]").GetAttributeValue("src", ""));
                    IWebBrowser webBrowserContent = this.RunPage.ShowWebPage(contentPageUrl, "pdf_content", 30000, false, WebBrowserType.Chromium, true);
                    while ((html == null || !html.Contains("name=\"plugin\"")) && waitingTimeout >= totalWaiting)
                    {
                        totalWaiting += interval;
                        Thread.Sleep(interval);
                        html = this.RunPage.InvokeGetPageHtml(webBrowserContent);
                    }
                    if (totalWaiting > waitingTimeout)
                    {
                        if (html.Contains("Sorry, we are unable to retrieve the document you requested."))
                        {
                            this.RunPage.InvokeAppendLogText("Sorry, we are unable to retrieve the document you requested. keywords = " + keywords + ", itemName = " + itemName + ", pageUrl = " + pageUrl, LogLevelType.System, true);
                        }
                        else
                        {
                            throw new Exception("页面加载失败_webBrowserContent, keywords = " + keywords + ", itemName = " + itemName + ", pageUrl = " + pageUrl);
                        }
                    }
                    else
                    {

                        HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                        htmlDoc.LoadHtml(html);
                        HtmlNode pdfFileUrlNode = htmlDoc.DocumentNode.SelectSingleNode("//embed[@name=\"plugin\"]");
                        string pdfFileUrl = CommonUtil.UrlDecodeSymbolAnd(pdfFileUrlNode.GetAttributeValue("src", ""));
                        byte[] fileBytes = this.RunPage.GetFileByRequest(pdfFileUrl, null, false, 1000, 1000 * 60 * 5, false, 1000);
                        this.RunPage.SaveFile(fileBytes, filePath);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        private string GetFilePath(string keywords, string itemName)
        {
            string sourceDirPath = this.RunPage.GetDetailSourceFileDir();
            if (!Directory.Exists(sourceDirPath))
            {
                Directory.CreateDirectory(sourceDirPath);
            }

            string dirPath = this.RunPage.GetFilePath(keywords, sourceDirPath);
            if (!Directory.Exists(dirPath))
            {
                Directory.CreateDirectory(dirPath);
            }

            string filePath = this.RunPage.GetFilePath(itemName, dirPath);
            return filePath;
        }

        private void SaveItemList(ExcelWriter ew, List<string> itemList)
        {
            for (int i = 0; i < itemList.Count; i++)
            {
                string itemName = itemList[i]; 
                Dictionary<string, string> row = new Dictionary<string, string>();
                row.Add("itemName", itemName); 
                ew.AddRow(row);
            }
        }

        private bool GotoNextPage(IWebBrowser webBrowser, int pageIndex, string keywords)
        {
            bool hasNext = true;
            string html = this.RunPage.InvokeGetPageHtml(webBrowser);
            HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
            htmlDoc.LoadHtml(html);
            HtmlNode nextBtnNode = htmlDoc.DocumentNode.SelectSingleNode("//ul[@class=\"results-paging nav-list\"]/li/a[@class=\"arrow-link legacy-link next\"]");
            if (nextBtnNode != null)
            {
                if (nextBtnNode.GetAttributeValue("disabled", "false") == "false")
                {
                    hasNext = true;
                }
                else
                {
                    hasNext = false;
                }
            }
            else
            {
                hasNext = false;
            }

            if (hasNext)
            {
                string nextBtnId = nextBtnNode.GetAttributeValue("id", "");

                string scriptCode = "document.getElementById('" + nextBtnId + "').click();";
                this.RunPage.InvokeAddScriptMethod(webBrowser, scriptCode);
                this.RunPage.ShowTabPage("EBSCOHost");
                Thread.Sleep(5000);
                int interval = 4000;
                int waitingTimeout = 30000;
                int totalWaiting = 0;
                html = "";
                while ((html == null ||!html.Contains("<li>" + pageIndex.ToString() + "</li>")) && waitingTimeout >= totalWaiting)
                {
                    totalWaiting += interval;
                    Thread.Sleep(interval);
                    html = this.RunPage.InvokeGetPageHtml(webBrowser);
                }
                if (totalWaiting > waitingTimeout)
                {
                    throw new Exception("打开下一页失败, pageIndex = " + pageIndex.ToString() + ", keywords = " + keywords);
                } 
            }
            return hasNext;
        }

        private void ExpandVolume(IWebBrowser webBrowser, string id)
        {
            string scriptCode = "var node = document.getElementById('" + id + "').click();";
            webBrowser.AddScriptMethod(scriptCode);
            Thread.Sleep(3000);
        } 

    }
}
