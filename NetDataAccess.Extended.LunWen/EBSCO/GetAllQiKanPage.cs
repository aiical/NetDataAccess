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
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace NetDataAccess.Extended.LunWen.EBSCO
{
    [System.Runtime.InteropServices.ComVisibleAttribute(true)]
    public class GetAllQiKanPage : ExternalRunWebPage
    {
        public override void GetDataByOtherAcessType(Dictionary<string, string> listRow)
        {
            string pageName = listRow[SysConfig.DetailPageNameFieldName];
            string pageUrl = listRow["pageUrl"];
            string keywords = listRow["keywords"];
            string moreKeywordStr = listRow["moreKeywords"];
            string[] morekeyWords = moreKeywordStr.Split(new string[] { ";" }, StringSplitOptions.RemoveEmptyEntries);
            List<string> moreKeyWordList = new List<string>(morekeyWords);

            IWebBrowser webBrowser = this.ShowEBSCOHostPage(pageUrl);
            this.ClickFullTextLink(webBrowser);

            this.DoSearch(webBrowser, keywords, moreKeyWordList);


            int pageIndex = 1;
            bool hasNextPage = true;


            String sourceDir = this.RunPage.GetDetailSourceFileDir();
            string sourceFilePath = this.RunPage.GetFilePath(pageUrl, sourceDir);
            ExcelWriter sourceEW = this.GetExcelWriter(sourceFilePath);

            while (hasNextPage)
            {
                List<string> itemList = this.GetListPageItems(keywords, webBrowser);
                this.SaveItemList(sourceEW, itemList);
                pageIndex++;
                hasNextPage = this.GotoNextPage(webBrowser, pageIndex, keywords);
            }
            sourceEW.SaveToDisk();
        } 

        public override bool AfterAllGrab(IListSheet listSheet)
        {
            string exportDir = this.RunPage.GetExportDir();
            string exportFilePath = Path.Combine(exportDir, "论文_ScienceDirect_论文列表页.xlsx");
            ExcelWriter resultWriter = this.GetExcelWriter(exportFilePath);

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> listRow = listSheet.GetRow(i);
                string pageUrl = listRow[SysConfig.DetailPageUrlFieldName];
                String sourceDir = this.RunPage.GetDetailSourceFileDir();
                string sourceFilePath = this.RunPage.GetFilePath(pageUrl, sourceDir);
                ExcelReader er = new ExcelReader(sourceFilePath);
                int rowCount = er.GetRowCount();
                for (int j = 0; j < rowCount; j++)
                {
                    Dictionary<string, string> resultRow = er.GetFieldValues(j);
                    resultWriter.AddRow(resultRow);
                }
            }
            resultWriter.SaveToDisk();
            return true;
        }

        private ExcelWriter GetExcelWriter(string filePath)        
        { 
            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("itemName", 0); 
            ExcelWriter resultEW = new ExcelWriter(filePath, "List", resultColumnDic, null);
            return resultEW;
        } 

        private IWebBrowser ShowEBSCOHostPage(string pageUrl)
        {
            try
            {
                string tabName = "EBSCOHost";
                IWebBrowser webBrowser = this.RunPage.ShowWebPage(pageUrl, tabName, 30000, false, WebBrowserType.Chromium, true); 

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

        private void ClickFullTextLink(IWebBrowser webBrowser)
        {
            string scriptCode = "document.getElementsByClassName('profileBodyBold')[1].click();";
            this.RunPage.InvokeAddScriptMethod(webBrowser, scriptCode);
            int interval = 2000;
            int waitingTimeout = 30000;
            int totalWaiting = 0;
            string html = "";
            while (!html.Contains("学术（同行评审）期刊") && waitingTimeout >= totalWaiting)
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

        private void DoSearch(IWebBrowser webBrowser, string keywords, List<string> moreKeywordList)
        {
            string scriptCode = "document.getElementById('common_SO').click();"
                                + "document.getElementById('common_SO').value='" + keywords + "';"
                                +"document.getElementById('common_FT').click();"
                                + "document.getElementById('common_RV').click();";
            this.RunPage.InvokeAddScriptMethod(webBrowser, scriptCode);
            int interval = 2000;
            int waitingTimeout = 30000;
            int totalWaiting = 0;
            string html = "";

            Thread.Sleep(interval);
            string scriptSubmitCode = "document.getElementById('ctl00_ctl00_MainContentArea_MainContentArea_ctrlLimiters_btnSearch').click()";
            this.RunPage.InvokeAddScriptMethod(webBrowser, scriptSubmitCode);

            while (!html.Contains("下一个") && waitingTimeout >= totalWaiting)
            {
                totalWaiting += interval;
                Thread.Sleep(interval);
                html = this.RunPage.InvokeGetPageHtml(webBrowser);
            }
            if (totalWaiting > waitingTimeout)
            {
                throw new Exception("跳转到搜索结果页面超时");
            }

            if (moreKeywordList != null && moreKeywordList.Count > 0)
            {

                string scriptACode = "document.getElementById('multiSelectCluster_JournalTrigger').click();"
                                     + "var moreLinkNodes = document.getElementsByClassName('panelShowMore evt-select-mulitple');"
                                     + "var targetLinkNode = null;"
                                     + "for(var i=0;i<moreLinkNodes.length;i++){"
                                     + "  var moreLinkNode = moreLinkNodes[i];" 
                                     + "  if(moreLinkNode.parentElement.parentElement.getAttribute('id') == 'multiSelectCluster_JournalContent'){"
                                     + "    targetLinkNode = moreLinkNode;"
                                     + "    break;"
                                     + "  }"
                                     + "}"
                                     + "alert(targetLinkNode.getAttribute('class'));"
                                     + "targetLinkNode.focus();";
                this.RunPage.InvokeAddScriptMethod(webBrowser, scriptACode);
                Thread.Sleep(10000);
                totalWaiting = 0;

                while (waitingTimeout >= totalWaiting)
                {
                    totalWaiting += interval;
                    Thread.Sleep(interval);
                    html = this.RunPage.InvokeGetPageHtml(webBrowser);
                    HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                    htmlDoc.LoadHtml(html);
                    HtmlNodeCollection jNodes = htmlDoc.DocumentNode.SelectNodes("//table[@class=\"limiter-table\"]/tbody/tr/td[@class=\"lim-select\"]/input");
                    HtmlNodeCollection jLabels = htmlDoc.DocumentNode.SelectNodes("//table[@class=\"limiter-table\"]/tbody/tr/td[@class=\"lim-name\"]/label");
                    if (jNodes != null)
                    {
                        for (int i = 0; i < jNodes.Count; i++)
                        {
                            HtmlNode jNode = jNodes[i];
                            HtmlNode jLabel = jLabels[i];
                            string jId = jNode.GetAttributeValue("id", "");
                            string jName = CommonUtil.HtmlDecode(jLabel.InnerText).Trim().ToLower();
                            if (moreKeywordList.Contains(jName))
                            {
                                string scriptJCode = "document.getElementById('" + jId + "').click();";
                                this.RunPage.InvokeAddScriptMethod(webBrowser, scriptJCode);
                            }
                        }
                        break;
                    }
                }
                if (totalWaiting > waitingTimeout)
                {
                    throw new Exception("弹出选择出版物窗口超时");
                }
                else
                {
                    string scriptRefreshCode = "document.getElementByClassName('button primary-action evt-update-btn').click();";
                    this.RunPage.InvokeAddScriptMethod(webBrowser, scriptRefreshCode);

                    totalWaiting = 0;
                    html = "";
                    while (!html.Contains("<h4 class=\"bb-heading\">出版物</h4>") && waitingTimeout >= totalWaiting)
                    {
                        totalWaiting += interval;
                        Thread.Sleep(interval);
                        html = this.RunPage.InvokeGetPageHtml(webBrowser);
                    }
                    if (totalWaiting > waitingTimeout)
                    {
                        throw new Exception("更新出版物后没有刷新出来界面");
                    }
                }                
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
                    this.DownloadBaseInfoPage(baseInfoPageUrl, keywords, itemName);
                }
                if (htmlPageUrl.Length > 0)
                {
                    this.DownloadHtmlPage(htmlPageUrl, keywords, itemName);
                }
                if (pdfPageUrl.Length > 0)
                {
                    this.DownloadPdfPage(pdfPageUrl, keywords, itemName);
                }
            }
            return itemName;
        }

        private void DownloadBaseInfoPage(string pageUrl, string keywords, string itemName)
        {
            string filePath = this.GetFilePath(keywords, itemName) + "_baseInfo";
            if (!File.Exists(filePath))
            {
                IWebBrowser webBrowser = this.RunPage.ShowWebPage(pageUrl, "baseInfo", 30000, false, WebBrowserType.Chromium, true);

                int interval = 2000;
                int waitingTimeout = 30000;
                int totalWaiting = 0;
                string html = "";
                while (!html.Contains("详细记录") && waitingTimeout >= totalWaiting)
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

                int interval = 2000;
                int waitingTimeout = 30000;
                int totalWaiting = 0;
                string html = "";
                while (!html.Contains("HTML 全文") && waitingTimeout >= totalWaiting)
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
            string filePath = this.GetFilePath(keywords, itemName) + "_html";
            if (!File.Exists(filePath))
            {
                IWebBrowser webBrowser = this.RunPage.ShowWebPage(pageUrl, "html", 30000, false, WebBrowserType.Chromium, true);

                int interval = 2000;
                int waitingTimeout = 30000;
                int totalWaiting = 0;
                string html = "";
                while (!html.Contains("plugin") && waitingTimeout >= totalWaiting)
                {
                    totalWaiting += interval;
                    Thread.Sleep(interval);
                    html = this.RunPage.InvokeGetPageHtml(webBrowser);
                }
                if (totalWaiting > waitingTimeout)
                {
                    throw new Exception("页面加载失败_html, keywords = " + keywords + ", itemName = " + itemName + ", pageUrl = " + pageUrl);
                }

                HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                htmlDoc.LoadHtml(html);
                HtmlNode pdfFileUrlNode = htmlDoc.DocumentNode.SelectSingleNode("//embed[@id=\"plugin\"]");
                string pdfFileUrl = pdfFileUrlNode.GetAttributeValue("href", "");
                byte[] fileBytes = this.RunPage.GetFileByRequest(pdfFileUrl, null, false, 1000, 30000, false, 1000);
                this.RunPage.SaveFile(fileBytes, filePath);
            }
        }


        private string GetFilePath(string keywords, string itemName)
        {
            string sourceDirPath = this.RunPage.GetDetailSourceFileDir();
            if (!Directory.Exists(sourceDirPath))
            {
                Directory.CreateDirectory(sourceDirPath);
            }

            string dirPath = Path.Combine(sourceDirPath, keywords);
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
                int interval = 2000;
                int waitingTimeout = 30000;
                int totalWaiting = 0;
                html = "";
                while (!html.Contains("<li>" + pageIndex.ToString() + "</li>") && waitingTimeout >= totalWaiting)
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
