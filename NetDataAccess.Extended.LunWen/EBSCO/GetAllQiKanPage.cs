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
using System.Drawing;
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
        public override void GetDataByOtherAccessType(Dictionary<string, string> listRow)
        {
            try
            {
                string pageMarkUrl = listRow[SysConfig.DetailPageUrlFieldName];
                string pageName = listRow[SysConfig.DetailPageNameFieldName];
                string pageUrl = listRow["pageUrl"];
                string keywords = listRow["keywords"];
                string moreKeywordStr = listRow["moreKeywords"];
                string[] morekeyWords = moreKeywordStr.Split(new string[] { ";" }, StringSplitOptions.RemoveEmptyEntries);
                List<string> moreKeyWordList = new List<string>(morekeyWords);

                IWebBrowser webBrowser = this.ShowEBSCOHostPage(pageUrl);
                this.ClickFullTextLink(webBrowser);


                String sourceDir = this.RunPage.GetDetailSourceFileDir();
                string sourceFilePath = this.RunPage.GetFilePath(pageMarkUrl, sourceDir);
                ExcelWriter sourceEW = this.GetExcelWriter(sourceFilePath);

               int fromYear = this.DoSearch(webBrowser, keywords, moreKeyWordList);

               for (int i = fromYear; i <= 2018; i++)
               {

                   string yearSourceFilePath = this.RunPage.GetFilePath(pageName, sourceDir) + "_" + i.ToString();
                   if (!File.Exists(yearSourceFilePath))
                   {

                       ExcelWriter yearSourceEW = this.GetExcelWriter(yearSourceFilePath);

                       this.GoToYear(i, webBrowser);

                       int pageIndex = 1;
                       bool hasNextPage = true;

                       while (hasNextPage)
                       {
                           List<string> itemList = this.GetListPageItems(keywords, webBrowser);
                           this.SaveItemList(yearSourceEW, itemList);
                           pageIndex++;
                           hasNextPage = this.GotoNextPage(webBrowser, pageIndex, keywords);
                       }
                       yearSourceEW.SaveToDisk();
                   }

                   ExcelReader yearEr = new ExcelReader(yearSourceFilePath);
                   int yearItemCount = yearEr.GetRowCount();
                   for (int j = 0; j < yearItemCount; j++)
                   {
                       Dictionary<string, string> yearRow = yearEr.GetFieldValues(j);
                       sourceEW.AddRow(yearRow);
                   }
                   yearEr.Close();
               }

                sourceEW.SaveToDisk();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void GoToYear(int year, IWebBrowser webBrowser)
        {
            string showSetYearWinScriptCode = "document.getElementById('ctl00_ctl00_Column1_Column1_ctl00_searchOptionsLink').click()";
            this.RunPage.InvokeAddScriptMethod(webBrowser, showSetYearWinScriptCode);
            Thread.Sleep(2000);

            int interval = 4000;
            int waitingTimeout = 30000;
            int totalWaiting = 0;
            string html = "";
            while ((html == null || !html.Contains("重新设置")) && waitingTimeout >= totalWaiting)
            {
                totalWaiting += interval;
                Thread.Sleep(interval);
                html = this.RunPage.InvokeGetPageHtml(webBrowser);
            }
            if (totalWaiting > waitingTimeout)
            {
                throw new Exception("打开筛选年份窗口失败, year = " + year.ToString());
            }

            string setScriptCode = "document.getElementsByName('common_DT1_FromYear')[1].value = " + year.ToString() + ";"
                                    + "document.getElementsByName('common_DT1_ToYear')[1].value = " + year.ToString() + ";";
            this.RunPage.InvokeAddScriptMethod(webBrowser, setScriptCode);
            Thread.Sleep(1000);

            string gotoYearScriptCode = "document.getElementById('ctrlLimiters_SubmitButtonTop').click()";
            this.RunPage.InvokeAddScriptMethod(webBrowser, gotoYearScriptCode);
            Thread.Sleep(2000);
            this.RunPage.ShowTabPage("EBSCOHost");

              
            totalWaiting = 0; 
            while ((html == null || !html.Contains(year.ToString() + "0101-" + year.ToString() + "1231")) && waitingTimeout >= totalWaiting)
            {
                totalWaiting += interval;
                Thread.Sleep(interval);
                html = this.RunPage.InvokeGetPageHtml(webBrowser);
            }
            if (totalWaiting > waitingTimeout)
            {
                throw new Exception("筛选年份失败, year = " + year.ToString());
            }
        }

        public List<string> GetDirFileNames(string dir)
        {
            string[] fileNames = Directory.GetFiles(dir);
            List<string> fileNameList = new List<string>();
            foreach (string fileName in fileNames)
            {
                fileNameList.Add(Path.GetFileName(fileName));
            }
            return fileNameList;
        }

        public override bool AfterAllGrab(IListSheet listSheet)
        {
            this.GetFileInfos(listSheet);
            this.GetKeywordsInfos(listSheet);
            return true;
        }
        private void GetFileInfos(IListSheet listSheet)
        {
            string exportDir = this.RunPage.GetExportDir();
            string exportFilePath = Path.Combine(exportDir, "论文_EBSCO_论文列表.xlsx");
            ExcelWriter resultWriter = this.GetResultExcelWriter(exportFilePath);
            String sourceDir = this.RunPage.GetDetailSourceFileDir();
            List<string> fileNameList = GetDirFileNames(sourceDir);

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> listRow = listSheet.GetRow(i);
                string publicationName = listRow[SysConfig.DetailPageNameFieldName];
                string moreKeywords = listRow["moreKeywords"];
                this.RunPage.InvokeAppendLogText("处理期刊, publicationName = " + publicationName + ", " + i.ToString() + "/" + listSheet.RowCount.ToString(), LogLevelType.System, true);
                foreach (string fileName in fileNameList)
                {
                    if (fileName.StartsWith(publicationName + "_"))
                    {
                        int year = int.Parse(fileName.Replace(publicationName + "_", ""));

                        string pubYearExcelFilePath = Path.Combine(sourceDir, fileName);
                        ExcelReader er = new ExcelReader(pubYearExcelFilePath);
                        int rowCount = er.GetRowCount();
                        for (int j = 0; j < rowCount; j++)
                        {
                            Dictionary<string, string> pubYearRow = er.GetFieldValues(j);
                            string itemName = pubYearRow["itemName"];
                            string pubDir = Path.Combine(sourceDir, publicationName);
                            string itemBaseInfoFilePath = this.RunPage.GetFilePath(itemName, pubDir) + "_baseInfo";
                            try
                            {
                                if (!File.Exists(itemBaseInfoFilePath))
                                {
                                    itemBaseInfoFilePath = Path.Combine(pubDir, CommonUtil.ProcessFileName(itemName, "_") + "_baseInfo");
                                    if (!File.Exists(itemBaseInfoFilePath))
                                    {
                                        throw new Exception("不存在文件， itemName =  " + itemName);
                                    }
                                }
                                Dictionary<string, object> fileRow = this.GetFileInfos(itemBaseInfoFilePath, moreKeywords.Length == 0 ? publicationName : moreKeywords);
                                fileRow.Add("publication", publicationName);
                                fileRow.Add("year", year);
                                resultWriter.AddRow(fileRow);
                            }
                            catch (Exception ex)
                            {
                                this.RunPage.InvokeAppendLogText("错误, " + ex.Message + " itemName = " + itemName, LogLevelType.Error, true);
                                //throw ex;
                            }
                        }
                    }
                }
            }
            resultWriter.SaveToDisk();
        }
        private void GetKeywordsInfos(IListSheet listSheet)
        {
            string exportDir = this.RunPage.GetExportDir();
            string exportFilePath = Path.Combine(exportDir, "论文_EBSCO_论文关键字.xlsx");
            ExcelWriter resultWriter = this.GetKeywordsResultExcelWriter(exportFilePath);
            String sourceDir = this.RunPage.GetDetailSourceFileDir();
            List<string> fileNameList = GetDirFileNames(sourceDir);

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> listRow = listSheet.GetRow(i);
                string publicationName = listRow[SysConfig.DetailPageNameFieldName];
                string moreKeywords = listRow["moreKeywords"];
                this.RunPage.InvokeAppendLogText("处理期刊关键字, publicationName = " + publicationName + ", " + i.ToString() + "/" + listSheet.RowCount.ToString(), LogLevelType.System, true);
                foreach (string fileName in fileNameList)
                {
                    if (fileName.StartsWith(publicationName + "_"))
                    {
                        int year = int.Parse(fileName.Replace(publicationName + "_", ""));

                        string pubYearExcelFilePath = Path.Combine(sourceDir, fileName);
                        ExcelReader er = new ExcelReader(pubYearExcelFilePath);
                        int rowCount = er.GetRowCount();
                        for (int j = 0; j < rowCount; j++)
                        {
                            Dictionary<string, string> pubYearRow = er.GetFieldValues(j);
                            string itemName = pubYearRow["itemName"];
                            string pubDir = Path.Combine(sourceDir, publicationName);
                            string itemBaseInfoFilePath = this.RunPage.GetFilePath(itemName, pubDir) + "_baseInfo";
                            try
                            {
                                if (!File.Exists(itemBaseInfoFilePath))
                                {
                                    itemBaseInfoFilePath = Path.Combine(pubDir, CommonUtil.ProcessFileName(itemName, "_") + "_baseInfo");
                                    if (!File.Exists(itemBaseInfoFilePath))
                                    {
                                        throw new Exception("不存在文件， itemName =  " + itemName);
                                    }
                                }
                                List<Dictionary<string, object>> fileRows = this.GetKeywordsInfos(itemBaseInfoFilePath, moreKeywords.Length == 0 ? publicationName : moreKeywords);
                                foreach (Dictionary<string, object> fileRow in fileRows)
                                {
                                    fileRow.Add("publication", publicationName);
                                    fileRow.Add("year", year);
                                    resultWriter.AddRow(fileRow);
                                }
                            }
                            catch (Exception ex)
                            {
                                this.RunPage.InvokeAppendLogText("错误, " + ex.Message + " itemName = " + itemName, LogLevelType.Error, true);
                                //throw ex;
                            }
                        }
                    }
                }
            }
            resultWriter.SaveToDisk();
        }

        private Dictionary<string, object> GetFileInfos(string filePath, string publicationNameStr)
        {
            string[] publicationNames = publicationNameStr.Split(new string[] { ";" }, StringSplitOptions.RemoveEmptyEntries);

            try
            {
                Dictionary<string, object> row = new Dictionary<string, object>();
                HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                htmlDoc.Load(filePath);
                HtmlNode fieldParentNode = htmlDoc.DocumentNode.SelectSingleNode("//dl[@id=\"citationFields\"]");
                HtmlNodeCollection dtNodes = fieldParentNode.SelectNodes("./dt");
                for (int i = 0; i < dtNodes.Count; i++)
                {
                    HtmlNode dtNode = dtNodes[i];
                    string dtName = CommonUtil.HtmlDecode(dtNode.InnerText).Trim();
                    if (dtName.Contains("替代标题"))
                    {
                        //不处理
                    }
                    else if (dtName.Contains("标题:"))
                    {
                        HtmlNode ddNode = this.FindDDNode(dtNode);
                        string title = CommonUtil.HtmlDecode(ddNode.InnerText).Trim();
                        row.Add("title", title);
                    }
                    else if (dtName.Contains("作者单位"))
                    {
                        //不处理
                    }
                    else if (dtName.Contains("作者提供的关键字") || dtName.Contains("关键字"))
                    {
                        HtmlNode ddNode = this.FindDDNode(dtNode);
                        List<string> keywordList = new List<string>();
                        foreach (HtmlNode childNode in ddNode.ChildNodes)
                        {
                            if (childNode is HtmlTextNode)
                            {
                                string keyword = CommonUtil.HtmlDecode(childNode.InnerText).Trim();
                                if (keyword.Length > 0)
                                {
                                    keywordList.Add(keyword);
                                }
                            }
                        }
                        row.Add("keywords", CommonUtil.StringArrayToString(keywordList.ToArray(), "; "));
                    }
                    else if (dtName.Contains("作者"))
                    {
                        HtmlNode ddNode = this.FindDDNode(dtNode);
                        List<string> authorList = new List<string>();
                        foreach (HtmlNode childNode in ddNode.ChildNodes)
                        {
                            if (childNode is HtmlTextNode)
                            {
                                string author = CommonUtil.HtmlDecode(childNode.InnerText).Trim();
                                if (author.Length > 0)
                                {
                                    authorList.Add(author);
                                }
                            }
                        }
                        row.Add("authors", CommonUtil.StringArrayToString(authorList.ToArray(), "; "));
                    }
                    else if (dtName.Contains("来源"))
                    {
                        HtmlNode ddNode = this.FindDDNode(dtNode);
                        string source = CommonUtil.HtmlDecode(ddNode.InnerText).Trim();
                        row.Add("source", source);

                        //是否来自于此期刊
                        bool matchedPublication = false;
                        foreach (string publicationName in publicationNames)
                        {
                            if (source.ToLower().StartsWith(publicationName.ToLower()))
                            {
                                matchedPublication = true;
                            }
                        }
                        row.Add("matchedPublication", matchedPublication ? "true" : "false");
                    }
                    else if (dtName.Contains("文献类型"))
                    {
                        HtmlNode ddNode = this.FindDDNode(dtNode);
                        string source = CommonUtil.HtmlDecode(ddNode.InnerText).Trim();
                        row.Add("documentType", source);
                    }
                    else if (dtName.Contains("摘要:") || dtName.Contains("摘要（英语）:"))
                    {
                        HtmlNode ddNode = this.FindDDNode(dtNode);
                        string abstracts = CommonUtil.HtmlDecode(ddNode.InnerText).Trim();
                        row.Add("abstract", abstracts);
                    }
                }
                return row;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private List<Dictionary<string, object>> GetKeywordsInfos(string filePath, string publicationNameStr)
        {
            string[] publicationNames = publicationNameStr.Split(new string[] { ";" }, StringSplitOptions.RemoveEmptyEntries);

            try
            {
                List<Dictionary<string, object>> rows = new List<Dictionary<string, object>>();
                HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                htmlDoc.Load(filePath);
                HtmlNode fieldParentNode = htmlDoc.DocumentNode.SelectSingleNode("//dl[@id=\"citationFields\"]");
                HtmlNodeCollection dtNodes = fieldParentNode.SelectNodes("./dt");

                for (int i = 0; i < dtNodes.Count; i++)
                {
                    HtmlNode dtNode = dtNodes[i];
                    string dtName = CommonUtil.HtmlDecode(dtNode.InnerText).Trim(); 
                    if (dtName.Contains("作者提供的关键字") || dtName.Contains("关键字"))
                    {
                        HtmlNode ddNode = this.FindDDNode(dtNode); 
                        foreach (HtmlNode childNode in ddNode.ChildNodes)
                        {
                            if (childNode is HtmlTextNode)
                            {
                                string keyword = CommonUtil.HtmlDecode(childNode.InnerText).Trim();

                                //单行keyword也可能带;号，需要处理
                                if (keyword.Length > 0)
                                {
                                    string[] kParts = keyword.Split(new string[] { ";"}, StringSplitOptions.RemoveEmptyEntries);
                                    foreach (string kPart in kParts)
                                    {
                                        Dictionary<string, object> row = new Dictionary<string, object>();
                                        row.Add("keyword", kPart.Trim());
                                        rows.Add(row);
                                    }
                                }
                            }
                        } 
                    } 
                }

                foreach (Dictionary<string, object> row in rows)
                {
                    try
                    {
                        for (int i = 0; i < dtNodes.Count; i++)
                        {
                            HtmlNode dtNode = dtNodes[i];
                            string dtName = CommonUtil.HtmlDecode(dtNode.InnerText).Trim();
                            if (dtName.Contains("作者提供的关键字") || dtName.Contains("关键字"))
                            {
                                //不处理
                            }
                            else 
                            if (dtName.Contains("替代标题"))
                            {
                                //不处理
                            }
                            else if (dtName.Contains("标题:"))
                            {
                                HtmlNode ddNode = this.FindDDNode(dtNode);
                                string title = CommonUtil.HtmlDecode(ddNode.InnerText).Trim();
                                row.Add("title", title);
                            }
                            else if (dtName.Contains("作者单位"))
                            {
                                //不处理
                            }
                            else if (dtName.Contains("作者"))
                            {
                                HtmlNode ddNode = this.FindDDNode(dtNode);
                                List<string> authorList = new List<string>();
                                foreach (HtmlNode childNode in ddNode.ChildNodes)
                                {
                                    if (childNode is HtmlTextNode)
                                    {
                                        string author = CommonUtil.HtmlDecode(childNode.InnerText).Trim();
                                        if (author.Length > 0)
                                        {
                                            authorList.Add(author);
                                        }
                                    }
                                }
                                row.Add("authors", CommonUtil.StringArrayToString(authorList.ToArray(), "; "));
                            }
                            else if (dtName.Contains("来源"))
                            {
                                HtmlNode ddNode = this.FindDDNode(dtNode);
                                string source = CommonUtil.HtmlDecode(ddNode.InnerText).Trim();
                                row.Add("source", source);

                                //是否来自于此期刊
                                bool matchedPublication = false;
                                foreach (string publicationName in publicationNames)
                                {
                                    if (source.ToLower().StartsWith(publicationName.ToLower()))
                                    {
                                        matchedPublication = true;
                                    }
                                }
                                row.Add("matchedPublication", matchedPublication ? "true" : "false");
                            }
                            else if (dtName.Contains("文献类型"))
                            {
                                HtmlNode ddNode = this.FindDDNode(dtNode);
                                string source = CommonUtil.HtmlDecode(ddNode.InnerText).Trim();
                                row.Add("documentType", source);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
                return rows;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        private HtmlNode FindDDNode(HtmlNode dtNode)
        {
            HtmlNode nextNode = dtNode.NextSibling;
            if (nextNode == null)
            {
                return null;
            }
            else if (nextNode.Name.ToLower() == "dd")
            {
                return nextNode;
            }
            else
            {
                return FindDDNode(nextNode);
            }
        }

        private ExcelWriter GetExcelWriter(string filePath)
        {
            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("itemName", 0);
            ExcelWriter resultEW = new ExcelWriter(filePath, "List", resultColumnDic, null);
            return resultEW;
        }

        private ExcelWriter GetResultExcelWriter(string filePath)
        {
            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("publication", 0);
            resultColumnDic.Add("year", 1);
            resultColumnDic.Add("title", 2);
            resultColumnDic.Add("authors", 3);
            resultColumnDic.Add("source", 4);
            resultColumnDic.Add("documentType", 5);
            resultColumnDic.Add("keywords", 6);
            resultColumnDic.Add("abstract", 7);
            resultColumnDic.Add("matchedPublication", 8);
            Dictionary<string, string> columnFormats = new Dictionary<string, string>();
            columnFormats.Add("year", "#0");
            ExcelWriter resultEW = new ExcelWriter(filePath, "List", resultColumnDic, columnFormats);
            return resultEW;
        }

        private ExcelWriter GetKeywordsResultExcelWriter(string filePath)
        {
            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("publication", 0);
            resultColumnDic.Add("year", 1);
            resultColumnDic.Add("title", 2);
            resultColumnDic.Add("authors", 3);
            resultColumnDic.Add("source", 4);
            resultColumnDic.Add("documentType", 5);
            resultColumnDic.Add("keyword", 6); 
            resultColumnDic.Add("matchedPublication", 7);
            Dictionary<string, string> columnFormats = new Dictionary<string, string>();
            columnFormats.Add("year", "#0");
            ExcelWriter resultEW = new ExcelWriter(filePath, "List", resultColumnDic, columnFormats);
            return resultEW;
        } 

        private IWebBrowser ShowEBSCOHostPage(string pageUrl)
        {
            try
            {
                string tabName = "EBSCOHost";
                IWebBrowser webBrowser = this.RunPage.ShowWebPage(pageUrl, tabName, 30000, false, WebBrowserType.Chromium, true);
                this.RunPage.ShowTabPage("EBSCOHost");

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
            int interval = 4000;
            int waitingTimeout = 30000;
            int totalWaiting = 0;
            string html = "";
            while ((html == null || !html.Contains("学术（同行评审）期刊")) && waitingTimeout >= totalWaiting)
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

        private int DoSearch(IWebBrowser webBrowser, string keywords, List<string> moreKeywordList)
        {
            string scriptCode = "document.getElementById('common_SO').click();"
                                + "document.getElementById('common_SO').value='" + keywords + "';"
                                + "document.getElementById('common_FT').click();"
                                + "document.getElementById('common_RV').click();";
            this.RunPage.InvokeAddScriptMethod(webBrowser, scriptCode);
            int interval = 4000;
            int waitingTimeout = 30000;
            int totalWaiting = 0;
            string html = "";

            Thread.Sleep(interval);
            string scriptSubmitCode = "document.getElementById('ctl00_ctl00_MainContentArea_MainContentArea_ctrlLimiters_btnSearch').click()";
            this.RunPage.InvokeAddScriptMethod(webBrowser, scriptSubmitCode);

            while ((html == null || !html.Contains("下一个")) && waitingTimeout >= totalWaiting)
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
                this.RunPage.InvokeScrollDocumentMethod(webBrowser, new Point(500, 500));
                Thread.Sleep(4000);

                string scriptACode = "document.getElementById('multiSelectCluster_JournalTrigger').click();";
                this.RunPage.InvokeAddScriptMethod(webBrowser, scriptACode);
                Thread.Sleep(4000);

                this.RunPage.InvokeScrollDocumentMethod(webBrowser, new Point(500, 700));
                Thread.Sleep(5000);
                totalWaiting = 0;

                html = this.RunPage.InvokeGetPageHtml(webBrowser);
                HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                htmlDoc.LoadHtml(html);
                HtmlNode aMoreNode = htmlDoc.DocumentNode.SelectSingleNode("//a[@id=\"multiSelectCluster_JournalTrigger\"]").ParentNode.SelectSingleNode("./div[@id=\"multiSelectCluster_JournalContent\"]/div[@class=\"controls\"]/a[@class=\"panelShowMore evt-select-mulitple\"]");
                if (aMoreNode == null)
                {
                    //只能是第一个关键词有效了
                    string moreKeyword = moreKeywordList[0];
                    HtmlNodeCollection jNodes = htmlDoc.DocumentNode.SelectNodes("//div[@id=\"multiSelectCluster_JournalContent\"]/ul/li/input");
                    HtmlNodeCollection jLabels = htmlDoc.DocumentNode.SelectNodes("//div[@id=\"multiSelectCluster_JournalContent\"]/ul/li/label");
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
                                break;
                            }
                        } 
                    }
                }
                else
                {
                    string scriptMoreCode = "var moreLinkNodes = document.getElementsByClassName('panelShowMore evt-select-mulitple');"
                                         + "var targetLinkNode = null;"
                                         + "for(var i=0;i<moreLinkNodes.length;i++){"
                                         + "  var moreLinkNode = moreLinkNodes[i];"
                                         + "  if(moreLinkNode.parentElement.parentElement.getAttribute('id') == 'multiSelectCluster_JournalContent'){"
                                         + "    targetLinkNode = moreLinkNode;"
                                         + "    break;"
                                         + "  }"
                                         + "}"
                                         + "targetLinkNode.click();";
                    this.RunPage.InvokeAddScriptMethod(webBrowser, scriptMoreCode);
                    Thread.Sleep(5000);
                    totalWaiting = 0;

                    while (waitingTimeout >= totalWaiting)
                    {
                        totalWaiting += interval;
                        Thread.Sleep(interval);
                        html = this.RunPage.InvokeGetPageHtml(webBrowser);
                        htmlDoc = new HtmlAgilityPack.HtmlDocument();
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

                    //string scriptRefreshCode = "document.getElementById('modalPanelForm').submit();";
                    string scriptRefreshCode = "document.getElementsByClassName('button primary-action evt-update-btn')[0].click();";
                    this.RunPage.InvokeAddScriptMethod(webBrowser, scriptRefreshCode);
                }

                totalWaiting = 0;
                html = "";
                while ((html == null || !html.Contains("<h4 class=\"bb-heading\">出版物</h4>")) && waitingTimeout >= totalWaiting)
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

            int fromYear = 2018;
            {
                html = this.RunPage.InvokeGetPageHtml(webBrowser);
                HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                htmlDoc.LoadHtml(html);
                HtmlNode fromYearNode = htmlDoc.DocumentNode.SelectSingleNode("//input[@id=\"ctl00_ctl00_Column1_Column1_ctl00_ctrlResultsDualSliderDate_txtFilterDateFrom\"]");
                string fromYearStr = fromYearNode.GetAttributeValue("value", "");
                fromYear = int.Parse(fromYearStr);
            }
            return fromYear;
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
