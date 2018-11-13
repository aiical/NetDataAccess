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

namespace NetDataAccess.Extended.LunWen.ScienceDirect
{
    [System.Runtime.InteropServices.ComVisibleAttribute(true)]
    public class GetAllQiKanListPage : ExternalRunWebPage
    {
        public override void GetDataByOtherAccessType(Dictionary<string, string> listRow)
        {
            string pageUrl = listRow[SysConfig.DetailPageUrlFieldName];
            int pageIndex = 0;
            bool hasNextPage = true;

            String sourceDir = this.RunPage.GetDetailSourceFileDir();
            string sourceFilePath = this.RunPage.GetFilePath(pageUrl, sourceDir);
            ExcelWriter sourceEW = this.GetExcelWriter(sourceFilePath);

            while (hasNextPage)
            {
                pageIndex++;
                IWebBrowser webBrowser = this.GetPaperListPageUrlsByWebBrowser(pageUrl, pageIndex);
                this.SavePaperListPageUrls(sourceEW, webBrowser);
                hasNextPage = this.CheckHasNextPage(webBrowser);
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
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            ExcelWriter resultEW = new ExcelWriter(filePath, "List", resultColumnDic, null);
            return resultEW;
        }

        private IWebBrowser GetPaperListPageUrlsByWebBrowser(string pageUrl, int pageIndex)
        {
            pageUrl = pageUrl + "?page=" + pageIndex.ToString();
            try
            {
                string tabName = "scienceDirect";
                IWebBrowser webBrowser = this.RunPage.ShowWebPage(pageUrl, tabName, 30000, false, WebBrowserType.Chromium, true); 

                if (!webBrowser.Loaded())
                {
                    throw new Exception("页面加载失败");
                }
                else
                {
                    while (!this.CheckAllVolumesExpanded(webBrowser))
                    {

                    }

                    return webBrowser;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("页面处理失败. pageUrl = " + pageUrl, ex);
            }
        }

        private void SavePaperListPageUrls(ExcelWriter ew, IWebBrowser webBrowser)
        {
            string html = this.RunPage.InvokeGetPageHtml(webBrowser);
            HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
            htmlDoc.LoadHtml(html);
            HtmlNodeCollection linkNodes = htmlDoc.DocumentNode.SelectNodes("//div[@class=\"issue-item u-margin-s-bottom\"]/a[@class=\"anchor text-m\"]");
            if (linkNodes != null)
            {
                for (int i = 0; i < linkNodes.Count; i++)
                {
                    HtmlNode linkNode = linkNodes[i];
                    string url = "http://58.194.172.12/rwt/109/https/P75YPLUUMNVXK5UDMWTGT6UFMN4C6Z5QNF" + linkNode.GetAttributeValue("href", "");
                    Dictionary<string, string> row = new Dictionary<string, string>();
                    row.Add(SysConfig.DetailPageUrlFieldName, url);
                    row.Add(SysConfig.DetailPageNameFieldName, url);
                    ew.AddRow(row);
                }
            }
        }

        private bool CheckHasNextPage(IWebBrowser webBrowser)
        {
            string html = this.RunPage.InvokeGetPageHtml(webBrowser);
            HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
            htmlDoc.LoadHtml(html);
            HtmlNode nextButton = htmlDoc.DocumentNode.SelectSingleNode("//button[@rel=\"next\"]");
            if (nextButton != null)
            {
                if (nextButton.GetAttributeValue("disabled", "false") == "false")
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        } 

        private void ExpandVolume(IWebBrowser webBrowser, string id)
        {
            string scriptCode = "var node = document.getElementById('" + id + "').click();";
            webBrowser.AddScriptMethod(scriptCode);
            Thread.Sleep(3000);
        }

        private bool CheckAllVolumesExpanded(IWebBrowser webBrowser)
        {
            bool allExpanded = true;
            string html = this.RunPage.InvokeGetPageHtml(webBrowser);
            HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
            htmlDoc.LoadHtml(html);
            HtmlNodeCollection buttonNodes = htmlDoc.DocumentNode.SelectNodes("//button[@class=\"accordion-panel-title u-padding-s-ver u-text-left text-l js-accordion-panel-title\"]");
            if (buttonNodes != null)
            {
                for (int i = 0; i < buttonNodes.Count; i++)
                {
                    HtmlNode buttonNode = buttonNodes[i];
                    bool expanded = buttonNode.GetAttributeValue("aria-expanded", true);
                    if (!expanded)
                    {
                        allExpanded = false;
                        string id = buttonNode.GetAttributeValue("id", "");
                        this.ExpandVolume(webBrowser, id);
                    }
                }
            }
            return allExpanded;
        }   

    }
}
