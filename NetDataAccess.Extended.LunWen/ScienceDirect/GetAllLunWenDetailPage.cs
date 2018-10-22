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
    public class GetAllLunWenDetailPage : ExternalRunWebPage
    {
        private string _CurrentTabName = "";
        public override void WebBrowserHtml_AfterDoNavigate(string pageUrl, Dictionary<string, string> listRow, string tabName)
        {
            this.RunPage.ShowTabPage(tabName);
            _CurrentTabName = tabName;
        }

        public override void CheckRequestCompleteFile(string webPageText, Dictionary<string, string> listRow)
        {
            if (webPageText.Contains("Download full text in PDF"))
            {
                if (webPageText.IndexOf("Download full text in PDF") == webPageText.LastIndexOf("Download full text in PDF"))
                {
                    throw new Exception("没有获取到下载pdf页面的Download full text in PDF链接");
                }
            }
            else if (webPageText.Contains("Loading..."))
            {
                throw new Exception("页面未加载完成，Loading...");
            }
            if (webPageText.Contains("pdfLink") && !webPageText.Contains("PdfDropDownMenu"))
            {
                IWebBrowser webBrowser = this.RunPage.GetWebBrowserByName(_CurrentTabName);
                string scriptCode = "document.getElementById('pdfLink').click();";
                this.RunPage.InvokeAddScriptMethod(webBrowser, scriptCode);
                throw new Exception("正在获取下载地址");
            }
        }

        public override bool AfterAllGrab(IListSheet listSheet)
        {
            string exportDir = this.RunPage.GetExportDir();
            string pdfUrlFilePath = Path.Combine(exportDir, "论文_ScienceDirect_论文PDF页.xlsx");
            ExcelWriter pdfUrlWriter = this.GetExcelWriter(pdfUrlFilePath);
             
            string baseInfoFilePath = Path.Combine(exportDir, "论文_ScienceDirect_论文基本信息.xlsx");
            ExcelWriter baseInfoWriter = this.GetExcelWriter(baseInfoFilePath);

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> listRow = listSheet.GetRow(i);
                string pageUrl = listRow[SysConfig.DetailPageUrlFieldName];
                string prefixUrl = this.GetUrlPrefix(pageUrl);
                String sourceDir = this.RunPage.GetDetailSourceFileDir();
                string sourceFilePath = this.RunPage.GetFilePath(pageUrl, sourceDir);

                HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);

                string publication = CommonUtil.HtmlDecode(htmlDoc.DocumentNode.SelectSingleNode("//a[@class=\"publication-title-link\"]").InnerText).Trim();
                string host = CommonUtil.HtmlDecode(htmlDoc.DocumentNode.SelectSingleNode("//div[@class=\"publication-volume u-text-center\"]/div[@class=\"text-xs\"]").InnerText).Trim();
                string title = CommonUtil.HtmlDecode(htmlDoc.DocumentNode.SelectSingleNode("//span[@class=\"title-text\"]").InnerText).Trim();
                HtmlNodeCollection authorNodes = htmlDoc.DocumentNode.SelectNodes("//a[@class=\"author size-m workspace-trigger\"]");
                List<string> authorList = new List<string>();
                for (int j = 0; j < authorNodes.Count; j++)
                {
                    authorList.Add(CommonUtil.HtmlDecode(authorNodes[j].InnerText).Trim());
                }
                string authors = CommonUtil.StringArrayToString(authorList.ToArray(), ", ");

                HtmlNode abstractNode = htmlDoc.DocumentNode.SelectSingleNode("//div[@id=\"abstracts\"]/div/div/p");
                string abstracts = abstractNode == null ? "" : CommonUtil.HtmlDecode(abstractNode.InnerText).Trim();

                List<string> referenceList = new List<string>();
                HtmlNode referenceNode = htmlDoc.DocumentNode.SelectSingleNode("//section[@clas=\"bibliography-sec\"]/dl[@class=\"references\"]");
                if (referenceNode != null)
                {
                    HtmlNodeCollection dtNodes = referenceNode.SelectNodes("./dt[@class=\"label\"]");
                    HtmlNodeCollection ddNodes = referenceNode.SelectNodes("./dd[@class=\"reference\"]");
                    for (int j = 0; j < dtNodes.Count; j++)
                    {
                        HtmlNode dtNode = dtNodes[j];
                        HtmlNode ddNode = ddNodes[j];
                        StringBuilder ss = new StringBuilder();
                        ss.Append("###");
                        ss.Append(CommonUtil.HtmlDecode(dtNode.InnerText).Trim());

                        HtmlNode contributionNode = ddNode.SelectSingleNode("./div[@class=\"contribution\"]");
                        string contribution = contributionNode == null ? "" : CommonUtil.HtmlDecode(contributionNode.InnerText).Trim();

                        HtmlNode refTitleNode = ddNode.SelectSingleNode("./div[@class=\"contribution\"]/string[@class=\"title\"]");
                        string refTitle = refTitleNode == null ? "" : CommonUtil.HtmlDecode(refTitleNode.InnerText).Trim();

                        string refAuthors = contribution.Replace(refTitle, "");

                        ss.Append("#author#:" + refAuthors);
                        ss.Append("#title#:" + refTitle);

                        HtmlNode hostNode = ddNode.SelectSingleNode("./div[@class=\"host\"]");
                        string refHost = hostNode == null ? "" : CommonUtil.HtmlDecode(hostNode.InnerText).Trim();
                        ss.Append("#host#:" + refHost);
                        referenceList.Add(ss.ToString());
                    }
                }
                string references = CommonUtil.StringArrayToString(referenceList.ToArray(), "\r\n");

                HtmlNode pdfDownloadNode = htmlDoc.DocumentNode.SelectSingleNode("//div[@class=\"PdfDropDownMenu\"]/a");
                if (pdfDownloadNode != null)
                {
                    string pdfUrl = prefixUrl + pdfDownloadNode.GetAttributeValue("href", "");
                    Dictionary<string, string> pdfUrlRow = new Dictionary<string, string>();
                    pdfUrlRow.Add("detailPageUrl", pdfUrl);
                    pdfUrlRow.Add("detailPageName", pdfUrl);
                    pdfUrlRow.Add("publication", publication);
                    pdfUrlRow.Add("host", host);
                    pdfUrlRow.Add("title", title);
                    pdfUrlRow.Add("authors", authors);
                    pdfUrlRow.Add("abstracts", abstracts);
                    pdfUrlRow.Add("references", references);
                    pdfUrlRow.Add("url", pageUrl);
                    pdfUrlWriter.AddRow(pdfUrlRow); 
                }

                /*
                HtmlNodeCollection allSpanNodes = htmlDoc.DocumentNode.SelectNodes("//span[@class=\"anchor-text\"]");
                if (allSpanNodes != null)
                {
                    for (int j = 0; j < allSpanNodes.Count; j++)
                    {
                        HtmlNode spanNode = allSpanNodes[j];
                        string spanText = CommonUtil.HtmlDecode(spanNode.InnerText).Trim();
                        if (spanText == "Download full text in PDF")
                        {

                            HtmlNode parentNode = spanNode.ParentNode;
                            if (parentNode.Name == "a")
                            {
                                string pdfUrl = prefixUrl + parentNode.GetAttributeValue("href", "");
                                Dictionary<string, string> pdfUrlRow = new Dictionary<string, string>();
                                pdfUrlRow.Add("detailPageUrl", pdfUrl);
                                pdfUrlRow.Add("detailPageName", pdfUrl);
                                pdfUrlRow.Add("publication", publication);
                                pdfUrlRow.Add("host", host);
                                pdfUrlRow.Add("title", title);
                                pdfUrlRow.Add("authors", authors);
                                pdfUrlRow.Add("abstracts", abstracts);
                                pdfUrlRow.Add("references", references);
                                pdfUrlRow.Add("url", pageUrl);
                                pdfUrlWriter.AddRow(pdfUrlRow);
                                break;
                            }
                        }
                    }
                }
                 * */

                Dictionary<string, string> baseInfoRow = new Dictionary<string, string>();
                baseInfoRow.Add("publication", publication);
                baseInfoRow.Add("host", host);
                baseInfoRow.Add("title", title);
                baseInfoRow.Add("authors", authors);
                baseInfoRow.Add("abstracts", abstracts);
                baseInfoRow.Add("references", references);
                baseInfoRow.Add("url", pageUrl);
                baseInfoWriter.AddRow(baseInfoRow);
            }

            pdfUrlWriter.SaveToDisk();
            baseInfoWriter.SaveToDisk();
            return true;
        }

        private string GetUrlPrefix(string listPageUrl)
        {
            int endIndex = listPageUrl.IndexOf("science") - 1;
            string prefix = listPageUrl.Substring(0, endIndex);
            return prefix;
        }

        private ExcelWriter GetExcelWriter(string filePath)
        {
            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>(); 
            resultColumnDic.Add("publication", 1);
            resultColumnDic.Add("host", 2);
            resultColumnDic.Add("title", 3);
            resultColumnDic.Add("authors", 4);
            resultColumnDic.Add("abstracts", 5);
            resultColumnDic.Add("references", 6);
            resultColumnDic.Add("url", 7);

            ExcelWriter resultEW = new ExcelWriter(filePath, "List", resultColumnDic, null);
            return resultEW;
        }

        private ExcelWriter GetDownloadPdfExcelWriter(string filePath)
        {
            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("publication", 5);
            resultColumnDic.Add("host", 6);
            resultColumnDic.Add("title", 7);
            resultColumnDic.Add("authors", 8);
            resultColumnDic.Add("abstracts", 9);
            resultColumnDic.Add("references", 10);
            resultColumnDic.Add("url", 11);

            ExcelWriter resultEW = new ExcelWriter(filePath, "List", resultColumnDic, null);
            return resultEW;
        } 
    }
}
