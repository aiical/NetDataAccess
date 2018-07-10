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
using System.Drawing;
using NetDataAccess.Base.Reader;
using NetDataAccess.Extended.Linkedin.Common;

namespace NetDataAccess.Extended.Linkedin.KeyWordsSearchAtGoogle
{
    /// <summary>
    /// GetAllListPage
    /// </summary>
    [System.Runtime.InteropServices.ComVisibleAttribute(true)]
    public class GetAllPages : ExternalRunWebPage
    {
        private string _LinkedinLoginPageUrl = null;
        private string LinkedinLoginPageUrl
        {
            get
            {
                return _LinkedinLoginPageUrl;
            }
            set
            {
                _LinkedinLoginPageUrl = value;
            }
        }

        private string _LinkedinLoginSucceedCheckUrl = null;
        private string LinkedinLoginSucceedCheckUrl
        {
            get
            {
                return _LinkedinLoginSucceedCheckUrl;
            }
            set
            {
                _LinkedinLoginSucceedCheckUrl = value;
            }
        }

        private string _GoogleUrlPrefix = null;
        private string GoogleUrlPrefix
        {
            get
            {
                return _GoogleUrlPrefix;
            }
            set
            {
                _GoogleUrlPrefix = value;
            }
        } 
        
        private string[]  _GoogleLinkedinItemPostfix=null;
        private string[] GoogleLinkedinItemPostfix
        {
            get
            {
                if (this._GoogleLinkedinItemPostfix == null)
                {
                    this._GoogleLinkedinItemPostfix = new string[] { 
                    "| 领英",
                    "| LinkedIn",
                    "- LinkedIn"
                    };
                }
                return this._GoogleLinkedinItemPostfix;
            }
        }

        private string GetRandomGoogleUrl()
        {
            return  this.GoogleUrlPrefix + "/?gws_rd=ssl#q=" + CommonUtil.UrlEncode(ProcessGooglePage.GetRandomSearchValue());
        }

        private void VisitRandomPage()
        {
            for (int i = 0; i < 2; i++)
            {
                string randomUrl = GetRandomGoogleUrl();
                this.RunPage.ShowWebPage(randomUrl, "randomPage", SysConfig.WebPageRequestTimeout, false);
            }
        }

        public override void BeforeGrabOne(string pageUrl, Dictionary<string, string> listRow, bool existLocalFile)
        {
            /*
            if (!existLocalFile)
            {
                VisitRandomPage();
            }*/
            ProcessThread.SleepRandom(5000, 8000);
        }

        public override void AfterGrabOne(string pageUrl, Dictionary<string, string> listRow, bool needReGrab, bool existLocalFile)
        {
            /*
            if (!needReGrab && !existLocalFile)
            {
                VisitRandomPage();
            }*/
        }

        public override bool BeforeAllGrab()
        {
            string[] ps = this.Parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            this.GoogleUrlPrefix = ps[0];
            this.LinkedinLoginPageUrl = ps[1];
            this.LinkedinLoginSucceedCheckUrl = ps[2];

            string excelFilePath = this.RunPage.ExcelFilePath;
            ExcelReader er = new ExcelReader(excelFilePath, "List");

            Dictionary<string, int> columnNameToIndex = CommonUtil.InitStringIndexDic(new string[]{
                "detailPageUrl",
                "detailPageName", 
                "cookie",  
                "grabStatus",   
                "giveUpGrab",  
                "loginName",  
                "loginPassword",  
                "keyWords",  
                "公司名称",  
                "公司名称关键词",  
                "所属领域",  
                "备注"});

            ExcelWriter ew = new ExcelWriter(excelFilePath, "List", columnNameToIndex);
            var listCount = er.GetRowCount();
            for (int i = 0; i < listCount; i++)
            {
                Dictionary<string, string> listRow = er.GetFieldValues(i);
                string keyWords = listRow["keyWords"];
                if (keyWords.Contains(" "))
                {
                    keyWords = "\"" + keyWords + "\"";
                }
                listRow["detailPageUrl"] = this.GoogleUrlPrefix + "/?gws_rd=ssl#q=" + CommonUtil.UrlEncode(keyWords + " inurl:cn.linkedin.com/in/ ");
                listRow["detailPageName"] = CommonUtil.UrlEncode(listRow["keyWords"]);

                ew.AddRow(listRow);
            }
            er.Close();
            ew.SaveToDisk();
            return true;
        }
        public void Test()
        {  
            Dictionary<string, int> columnNameToIndex = CommonUtil.InitStringIndexDic(new string[]{
                "word"});

            ExcelWriter ew = new ExcelWriter("f:\\c.xlsx", "List", columnNameToIndex);
            string[] ssArray = new string[] { 
                "sina.com.cn",
                "xinhua.com",
                "twitter.com",
                "amazon.com",
                "baidu.com",
                "nytimes.com",
                "jd.com",
                "tmall.com",
                "sohu.com",
                "qq.com",
                "taobao.com",
                "tianya.com",
                "bustbuy.com"
            };
            var listCount = ssArray.Length;
            for (int i = 1; i < listCount; i++)
            {
                string word = ssArray[i];
                string[] ws = word.Split(new string[] { "\t" }, StringSplitOptions.RemoveEmptyEntries);
                Dictionary<string, string> listRow = new Dictionary<string, string>();
                listRow["word"] = ws[0];
                ew.AddRow(listRow);
            }
            ew.SaveToDisk(); 
        }

        public void AutoScroll(IRunWebPage runPage, WebBrowser webBrowser, int toPos, int maxStepLength, int minStepSleep, int maxStepSleep)
        {
            int pos = 0;
            Random random = new Random(DateTime.Now.Millisecond);
            while (pos < toPos)
            {
                int randomValue = random.Next(maxStepLength);
                pos += randomValue;
                runPage.InvokeScrollDocumentMethod(webBrowser, new Point(pos, pos));
                ProcessThread.SleepRandom(minStepSleep, maxStepSleep);
            }
        }

        public override void WebBrowserHtml_AfterPageLoaded(string pageUrl, Dictionary<string, string> listRow, WebBrowser webBrowser)
        {
            ProcessWebBrowser.AutoScroll(this.RunPage, webBrowser, 3000, 500, 1000, 2000);
            if (this.RunPage.InvokeCheckWebBrowserContains(webBrowser, new string[] { "系统检测到您的计算机网络中存在异常流量" }, true))
            {
                throw new Exception("Google系统检测到您的计算机网络中存在异常流量");
            }
        }

        public override bool AfterAllGrab(IListSheet listSheet)
        {
            var seedCount = listSheet.RowCount;
            for (int i = 0; i < seedCount; i++)
            {
                Dictionary<string, string> seedRow = listSheet.GetRow(i);
                try
                {
                    this.GetOneKeyWordsRelatedInfos(seedRow);
                }
                catch (Exception ex)
                {
                    this.RunPage.InvokeAppendLogText("使用关键词" + this.GetKeyWords(seedRow) + "爬取时出错. " + ex.Message, LogLevelType.System, true);
                }
            }
            return true;
        }

        private void GetOneKeyWordsRelatedInfos(Dictionary<string, string> seedRow)
        {

            //下一步必须执行
            bool isNewDo = false;
            string localLogFileName = null;
            string keyWords = this.GetKeyWords(seedRow);

            List<string> allListPageUrls = null;
            localLogFileName = "_" + this.GetLoginName(seedRow) + "_" + this.GetKeyWords(seedRow) + "_listPageUrl";
            if (SysConfig.SysExecuteType == SysExecuteType.Produce)
            {
                //如果是生产环境，那么直接爬取列表页
                allListPageUrls = this.GetAllListPages(this.GetSeedPageUrl(seedRow), keyWords);
                this.RunPage.SaveInfoToMiddleFile(localLogFileName, "listPageUrl", allListPageUrls);
                isNewDo = true;
            }
            else
            {
                //读取历史爬取的列表页地址文件 
                allListPageUrls = this.RunPage.TryGetInfoFromMiddleFile(localLogFileName, "listPageUrl");
                if (allListPageUrls == null)
                {
                    allListPageUrls = this.GetAllListPages(this.GetSeedPageUrl(seedRow), keyWords);
                    this.RunPage.SaveInfoToMiddleFile(localLogFileName, "listPageUrl", allListPageUrls);
                    isNewDo = true;
                }

            }

            List<Dictionary<string, string>> allPersonPageUrlInfos = null;
            localLogFileName = "_" + this.GetLoginName(seedRow) + "_" + this.GetKeyWords(seedRow) + "_personPageUrlInfo";
            if (SysConfig.SysExecuteType == SysExecuteType.Produce || isNewDo)
            {
                //如果是生产环境，那么直接解析列表页
                allPersonPageUrlInfos = this.GetPersonPageUrlsFromListPages(this.RunPage.GetDetailSourceFileDir(), allListPageUrls);
                this.RunPage.SaveInfoToMiddleFile(localLogFileName, new string[] { "personUrl", "personName" }, allPersonPageUrlInfos);
            }
            else
            {
                //读取历史解析获得的个人网页地址
                allPersonPageUrlInfos = this.RunPage.TryGetInfoFromMiddleFile(localLogFileName, new string[] { "personUrl", "personName" });
                if (allPersonPageUrlInfos == null)
                {
                    allPersonPageUrlInfos = this.GetPersonPageUrlsFromListPages(this.RunPage.GetDetailSourceFileDir(), allListPageUrls);
                    this.RunPage.SaveInfoToMiddleFile(localLogFileName, new string[] { "personUrl", "personName" }, allPersonPageUrlInfos);
                    isNewDo = true;
                }
            }
             
            LoginLinkedin.LoginByRandomUser(this.RunPage, this.LinkedinLoginPageUrl, this.LinkedinLoginSucceedCheckUrl);

            List<Dictionary<string,string>> allPersonPageInfosWithJustDownloadMark = null;
            localLogFileName = "_" + this.GetLoginName(seedRow) + "_" + this.GetKeyWords(seedRow) + "_personPageUrl";
            if (SysConfig.SysExecuteType == SysExecuteType.Produce || isNewDo)
            {
                //如果是生产环境，那么直接爬取个人详情页
                allPersonPageInfosWithJustDownloadMark = ProcessPersonPage.GetAllPersonPages(this.RunPage, allPersonPageUrlInfos, this.GetLoginName(seedRow), this.GetLoginPassword(seedRow));
                this.RunPage.SaveInfoToMiddleFile(localLogFileName, new string[]{ "personUrl", "personUrl","isJustDownload"}, allPersonPageInfosWithJustDownloadMark);
                isNewDo = true;
            }
            else
            {
                //读取历史生成的个人网页网址
                allPersonPageInfosWithJustDownloadMark = this.RunPage.TryGetInfoFromMiddleFile(this.GetLoginName(seedRow) + "." + this.GetKeyWords(seedRow) + ".personPageUrl", new string[] { "personUrl", "personUrl", "isJustDownload" });
                if (allPersonPageInfosWithJustDownloadMark == null)
                {
                    allPersonPageInfosWithJustDownloadMark = ProcessPersonPage.GetAllPersonPages(this.RunPage, allPersonPageUrlInfos, this.GetLoginName(seedRow), this.GetLoginPassword(seedRow));
                    this.RunPage.SaveInfoToMiddleFile(localLogFileName, new string[] { "personUrl", "personUrl", "isJustDownload" }, allPersonPageInfosWithJustDownloadMark);
                    isNewDo = true;
                }
            }

            List<Dictionary<string, string>> personInfoList = ProcessPersonPage.GetPersonInfoFromLocalPages(this.RunPage, allPersonPageInfosWithJustDownloadMark, true, this.GetKeyWords(seedRow));
             
            string personInfosResultFilePath = this.RunPage.GetFilePath("SearchResult_Google2Linkedin_" + this.GetLoginName(seedRow) + "_" + this.GetKeyWords(seedRow) + ".xlsx", this.RunPage.GetExportDir());

            ProcessPersonPage.SavePersonInfoToFile(this.RunPage, personInfoList, personInfosResultFilePath); 
        } 

        private string GetKeyWords(Dictionary<string, string> listRow)
        {
            return listRow["keyWords"];
        }

        private string GetLoginName(Dictionary<string, string> listRow)
        {
            return listRow["loginName"];
        }

        private string GetLoginPassword(Dictionary<string, string> listRow)
        {
            return listRow["loginPassword"];
        }

        private string GetSeedPageUrl(Dictionary<string, string> listRow)
        {
            return listRow[SysConfig.DetailPageUrlFieldName];
        }

        /// <summary>
        /// 循环获取所有列表页
        /// </summary>
        /// <param name="nextPageUrl"></param>
        private void GetListPageOneByOne(string seedPageUrl, string keyWords, List<string> allListPageUrls)
        {
            int pageIndex = 0;
            bool hasNextPage = true;
            while (hasNextPage)
            { 
                //ProcessWebBrowser.ClearWebBrowserTracks();
                string nextPageUrl = seedPageUrl + "&start=" + pageIndex * 10;
                hasNextPage = this.GetCurrentPageAndNextPageUrl(seedPageUrl, keyWords, nextPageUrl, allListPageUrls);
                pageIndex++;
            }
        }

        /// <summary>
        /// 获取当前列表页及下一页地址
        /// </summary>
        /// <param name="listPageUrl"></param>
        /// <returns></returns>
        private bool GetCurrentPageAndNextPageUrl(string seedPageUrl, string keyWords, string listPageUrl, List<string> allListPageUrls)
        {
            VisitRandomPage();

            string localFilePath = this.RunPage.GetFilePath(listPageUrl, this.RunPage.GetDetailSourceFileDir());

            if (!File.Exists(localFilePath))
            {
                string tabName = "ListPage";
                WebBrowser webBrowser = this.RunPage.ShowWebPage(listPageUrl, tabName, SysConfig.WebPageRequestTimeout, false); 
                try
                {
                    this.RunPage.CheckWebBrowserContainsForComplete(webBrowser, new string[] { keyWords }, SysConfig.WebPageRequestTimeout, true);
                }
                catch (Exception ex)
                {
                    string limitAlert = "计算机网络中存在异常流量";
                    string errorPageHtml = this.RunPage.InvokeGetPageHtml(tabName);
                    if (errorPageHtml.Contains(limitAlert))
                    {
                        ProcessWebBrowser.ClearWebBrowserTracks();
                        ProcessWebBrowser.ClearWebBrowserCookie();
                        this.RunPage.InvokeAppendLogText(limitAlert + ". 正在清理缓存, 并等待重新启动爬取.", LogLevelType.System, true);
                        throw new Exception("Google" + limitAlert);
                    }
                    else
                    {
                        throw ex;
                    }
                }

                string listPageHtml = this.RunPage.InvokeGetPageHtml(tabName);
                ProcessWebBrowser.AutoScroll(this.RunPage, webBrowser, 2000, 1000, 1000, 2000);
                this.RunPage.SaveFile(listPageHtml, localFilePath, Encoding.UTF8);

                allListPageUrls.Add(listPageUrl);


                string scriptMethodCode = "function myGetNextPageUrl(){"
                    + "var nextA = document.getElementById('pnnext');"
                    + "if(nextA == null){"
                    + "return '';"
                    + "}"
                    + "else{"
                    + "return nextA.getAttribute('href');"
                    + "}"
                    + "}";

                this.RunPage.InvokeAddScriptMethod(webBrowser, scriptMethodCode, this);
                string nextPageUrl = CommonUtil.UrlDecodeSymbolAnd((string)this.RunPage.InvokeDoScriptMethod(webBrowser, "myGetNextPageUrl", null));

                if (nextPageUrl != null && nextPageUrl.Length > 0)
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
                allListPageUrls.Add(listPageUrl);
                return true;
            }
        }

        /// <summary>
        /// 获取所有列表页内容
        /// </summary>
        /// <param name="listSheet"></param>
        private List<string> GetAllListPages(string seedPageUrl, string keyWords)
        { 
            try
            {
                List<string> allListPageUrls = new List<string>();
                string pageDir = this.RunPage.GetDetailSourceFileDir();
                string localFilePath = this.RunPage.GetFilePath(seedPageUrl, pageDir);
                StreamReader tr = new StreamReader(localFilePath, Encoding.UTF8);
                string webPageHtml = tr.ReadToEnd();

                HtmlAgilityPack.HtmlDocument pageHtmlDoc = new HtmlAgilityPack.HtmlDocument();
                pageHtmlDoc.LoadHtml(webPageHtml);

                HtmlNodeCollection liNodes = pageHtmlDoc.DocumentNode.SelectNodes("//div[@class=\"rc\"]");
                if (liNodes.Count == 0)
                {
                    throw new Exception("没有找到符合条件的人");
                }
                else
                {
                    HtmlNodeCollection pageNodes = pageHtmlDoc.DocumentNode.SelectNodes("//table[@id=\"nav\"]/tbody/tr/td/a[@class=\"fl\"]");

                    if (pageNodes.Count == 0)
                    {
                        allListPageUrls.Add(seedPageUrl);
                    }
                    else
                    {
                        this.GetListPageOneByOne(seedPageUrl, keyWords, allListPageUrls);
                    }
                }
                return allListPageUrls;
            }
            catch (Exception ex)
            {
                this.RunPage.InvokeAppendLogText(ex.Message, LogLevelType.Error, true);
                throw new Exception("抓取Google列表页失败", ex);
            }
        }

        /// <summary>
        /// 获取列表页首页地址
        /// </summary>
        /// <param name="liNodes"></param>
        /// <returns></returns>
        private string GetFirstListPageUrl(HtmlAgilityPack.HtmlDocument htmlDoc)
        {
            HtmlNodeCollection pageNodes = htmlDoc.DocumentNode.SelectNodes("//table[@id=\"nav\"]/tbody/tr/td/a[@class=\"fl\"]");

            HtmlNode pageNode = pageNodes[0];
            string fullUrl = CommonUtil.UrlDecode(CommonUtil.UrlDecodeSymbolAnd(pageNode.GetAttributeValue("href", "")));
            string[] urlSplits = fullUrl.Split(new string[] { "?" }, StringSplitOptions.RemoveEmptyEntries);
            string[] parameters = urlSplits[1].Split(new string[] { "&" }, StringSplitOptions.RemoveEmptyEntries);
            StringBuilder firstPageUrl = new StringBuilder();
            firstPageUrl.Append(urlSplits[0]);
            firstPageUrl.Append("?");
            for (int i = 0; i < parameters.Length; i++)
            {
                string p = parameters[i];
                if (i > 0)
                {
                    firstPageUrl.Append("&");
                }
                if (p.StartsWith("start="))
                {
                    firstPageUrl.Append("start=0");
                }
                else
                {
                    firstPageUrl.Append(p);
                }
            }
            return this.GoogleUrlPrefix + firstPageUrl.ToString();
        }

        /// <summary>
        /// 从页面地址中获取页码
        /// </summary>
        /// <param name="pageUrl"></param>
        /// <returns></returns>
        private string GetPageNum(string pageUrl)
        {
            string[] urlSplits = pageUrl.Split(new string[] { "?" }, StringSplitOptions.RemoveEmptyEntries);
            string[] parameters = urlSplits[1].Split(new string[] { "&" }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < parameters.Length; i++)
            {
                string p = parameters[i];
                if (p.StartsWith("start="))
                {
                    return p.Replace("start=", "");
                }
            }
            return "";
        }

        /// <summary>
        /// 从所有列表页中获取个人页面网址
        /// </summary>
        /// <param name="localDir"></param>
        /// <param name="pageUrls"></param>
        private List<Dictionary<string, string>> GetPersonPageUrlsFromListPages(string localDir, List<string> listPageUrls)
        {
            try
            {
                List<Dictionary<string, string>> allPersonPageUrlInfos = new List<Dictionary<string, string>>();
                foreach (string listPageUrl in listPageUrls)
                {
                    this.GetPersonPageUrls(localDir, listPageUrl, allPersonPageUrlInfos);
                }
                return allPersonPageUrlInfos;
            }
            catch (Exception ex)
            {
                this.RunPage.InvokeAppendLogText(ex.Message, LogLevelType.Error, true);
                throw new Exception("解析列表页获取个人页面地址失败", ex);
            }
        }

        private void GetPersonPageUrls(string localDir, string listPageUrl, List<Dictionary<string, string>> allPersonPageUrlInfos)
        {
            try
            {
                string listPageLocalPath = this.RunPage.GetFilePath(listPageUrl, localDir);
                HtmlAgilityPack.HtmlDocument pageHtmlDoc = HtmlDocumentHelper.Load(listPageLocalPath);
                HtmlNodeCollection allANodes = pageHtmlDoc.DocumentNode.SelectNodes("//div[@class=\"rc\"]/h3/a");
                foreach (HtmlNode aNode in allANodes)
                {
                    string personUrl = HtmlDocumentHelper.TryGetNodeAttributeValue(aNode, "data-href", true, true, null, null);
                    if (personUrl == null)
                    {
                        personUrl = HtmlDocumentHelper.TryGetNodeAttributeValue(aNode, "href", true, true, null, null);
                    }
                    if (personUrl.Contains(".linkedin.com/in/"))
                    {
                        try
                        {
                            string personName = aNode.InnerText.Trim();
                            foreach (string postfix in this.GoogleLinkedinItemPostfix)
                            {
                                personName = personName.Replace(postfix, "");
                            } 
                            Dictionary<string, string> personPageUrlInfo = new Dictionary<string, string>();
                            personUrl = CommonUtil.UrlDecode(personUrl);
                            personPageUrlInfo.Add("personUrl", personUrl);
                            personPageUrlInfo.Add("personName", personName.Trim());
                            allPersonPageUrlInfos.Add(personPageUrlInfo);
                        }
                        catch (Exception ex)
                        {
                            throw new Exception("获取个人网页地址时出错", ex);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                this.RunPage.InvokeAppendLogText(ex.Message, LogLevelType.Error, true);
                throw new Exception("解析Google列表页出错, listPageUrl = +" + listPageUrl, ex);
            }
        }
    }
}