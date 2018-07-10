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

namespace NetDataAccess.Extended.Linkedin.KeyWordsSearchAtBaiduHttpRequest
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

        private string _BaiduUrlPrefix = null;
        private string BaiduUrlPrefix
        {
            get
            {
                return _BaiduUrlPrefix;
            }
            set
            {
                _BaiduUrlPrefix = value;
            }
        } 
        
        private string[]  _BaiduLinkedinItemPostfix=null;
        private string[] BaiduLinkedinItemPostfix
        {
            get
            {
                if (this._BaiduLinkedinItemPostfix == null)
                {
                    this._BaiduLinkedinItemPostfix = new string[] { 
                    "| 领英",
                    "| LinkedIn",
                    "- LinkedIn"
                    };
                }
                return this._BaiduLinkedinItemPostfix;
            }
        }
        
        private string[]  _UserAgents=null;
        private string[] UserAgents
        {
            get
            {
                if (this._UserAgents == null)
                {
                    this._UserAgents = new string[] { 
                        "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:23.0) Gecko/20130406 Firefox/23.0",
                        "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:18.0) Gecko/20100101 Firefox/18.0",
                        "Mozilla/5.0 (Windows; U; Windows NT 6.1; en-US) AppleWebKit/533+ (KHTML, like Gecko) Element Browser 5.0", 
                        "IBM WebExplorer /v0.94', 'Galaxy/1.0 [en] (Mac OS X 10.5.6; U; en)", 
                        "Mozilla/5.0 (compatible; MSIE 10.0; Windows NT 6.1; WOW64; Trident/6.0)", 
                        "Opera/9.80 (Windows NT 6.0) Presto/2.12.388 Version/12.14", 
                        "Mozilla/5.0 (iPad; CPU OS 6_0 like Mac OS X) AppleWebKit/536.26 (KHTML, like Gecko) Version/6.0 Mobile/10A5355d Safari/8536.25", 
                        "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/28.0.1468.0 Safari/537.36", 
                        "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.0; Trident/5.0; TheWorld)"
                    };
                }
                return this._UserAgents;
            }
        }

        private string GetRandomAgent()
        {
            Random random = new Random(DateTime.Now.Millisecond);
            int rValue = random.Next(0, this.UserAgents.Length);
            return this.UserAgents[rValue];
        }

        public override void WebRequestHtml_BeforeSendRequest(string pageUrl, Dictionary<string, string> listRow, Base.Web.NDAWebClient client)
        {
            string userAgent = this.GetRandomAgent();
            client.Headers.Set("User-Agent", userAgent);
        } 

        public override bool BeforeAllGrab()
        {
            string[] ps = this.Parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            this.BaiduUrlPrefix = ps[0];
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
                string keyWords = "\"" + listRow["keyWords"] + "\"";
                string queryWords = CommonUtil.UrlEncode("site:(cn.linkedin.com) " + keyWords).Replace("+", "%20");
                listRow["detailPageUrl"] = this.BaiduUrlPrefix + "/s?wd=" + queryWords + "&oq=" + queryWords + "&tn=baiduadv";
                //listRow["detailPageUrl"] = this.BaiduUrlPrefix + "/s?wd=" + CommonUtil.UrlEncode("site:(cn.linkedin.com) " + keyWords);
                listRow["detailPageName"] =  CommonUtil.UrlEncode(listRow["keyWords"]);

                ew.AddRow(listRow);
            }
            er.Close();
            ew.SaveToDisk(); 
            return true;
        }

        private void GetPageFromBaidu(string pageUrl, string keyWords)
        {
            try
            {
                Proj_CompleteCheckList completeCheckList = new Proj_CompleteCheckList();
                Proj_CompleteCheck checkSiteName = new Proj_CompleteCheck();
                checkSiteName.CheckType = DocumentCompleteCheckType.TextExist;
                checkSiteName.CheckValue = "cn.linkedin.com";
                completeCheckList.Add(checkSiteName);
                Proj_CompleteCheck checkKeyWord = new Proj_CompleteCheck();
                checkKeyWord.CheckType = DocumentCompleteCheckType.TextExist;
                checkKeyWord.CheckValue = keyWords;
                completeCheckList.Add(checkKeyWord);

                string localPageFilePath = this.RunPage.GetFilePath(pageUrl, this.RunPage.GetDetailSourceFileDir());
                if (!File.Exists(localPageFilePath))
                {
                    string responseString = null;
                    try
                    {
                        responseString = this.RunPage.GetTextByRequest(pageUrl, null, false, 3000, SysConfig.WebPageRequestTimeout, Encoding.UTF8, null, null, true, Proj_DataAccessType.OtherAccessType, completeCheckList, 1000);
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                    this.RunPage.SaveFile(responseString, localPageFilePath, Encoding.UTF8);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public override bool CheckNeedGrab(Dictionary<string, string> listRow, string localPagePath)
        {
            string personInfosResultFilePath = this.RunPage.GetFilePath("_SearchResult_Baidu2Linkedin_" + this.GetKeyWords(listRow) + ".xlsx", this.RunPage.GetExportDir());
            return !File.Exists(personInfosResultFilePath);
        }

        public override void AfterGrabOne(string pageUrl, Dictionary<string, string> seedRow, bool needReGrab, bool existLocalFile)
        {
            try
            {
                this.GetPageFromBaidu(pageUrl, this.GetKeyWords(seedRow));

                //下一步必须执行
                bool isNewDo = false;
                string localLogFileName = null;

                List<string> allListPageUrls = null;
                localLogFileName = "_" + this.GetKeyWords(seedRow) + "_listPageUrl";
                if (SysConfig.SysExecuteType == SysExecuteType.Produce)
                {
                    //如果是生产环境，那么直接爬取列表页
                    allListPageUrls = this.GetAllListPages(this.GetSeedPageUrl(seedRow), this.GetKeyWords(seedRow));
                    this.RunPage.SaveInfoToMiddleFile(localLogFileName, "listPageUrl", allListPageUrls);
                    isNewDo = true;
                }
                else
                {
                    //读取历史爬取的列表页地址文件 
                    allListPageUrls = this.RunPage.TryGetInfoFromMiddleFile(localLogFileName, "listPageUrl");
                    if (allListPageUrls == null)
                    {
                        allListPageUrls = this.GetAllListPages(this.GetSeedPageUrl(seedRow), this.GetKeyWords(seedRow));
                        this.RunPage.SaveInfoToMiddleFile(localLogFileName, "listPageUrl", allListPageUrls);
                        isNewDo = true;
                    }

                }

                List<Dictionary<string, string>> allPersonPageUrlInfos = null;
                localLogFileName = "_" + this.GetKeyWords(seedRow) + "_personPageUrlInfo";
                if (SysConfig.SysExecuteType == SysExecuteType.Produce || isNewDo)
                {
                    //如果是生产环境，那么直接解析列表页
                    allPersonPageUrlInfos = this.GetPersonPageUrlsFromListPages(this.RunPage.GetDetailSourceFileDir(), allListPageUrls, this.GetKeyWords(seedRow));
                    this.RunPage.SaveInfoToMiddleFile(localLogFileName, new string[] { "personUrl", "personName" }, allPersonPageUrlInfos);
                }
                else
                {
                    //读取历史解析获得的个人网页地址
                    allPersonPageUrlInfos = this.RunPage.TryGetInfoFromMiddleFile(localLogFileName, new string[] { "personUrl", "personName" });
                    if (allPersonPageUrlInfos == null)
                    {
                        allPersonPageUrlInfos = this.GetPersonPageUrlsFromListPages(this.RunPage.GetDetailSourceFileDir(), allListPageUrls, this.GetKeyWords(seedRow));
                        this.RunPage.SaveInfoToMiddleFile(localLogFileName, new string[] { "personUrl", "personName" }, allPersonPageUrlInfos);
                        isNewDo = true;
                    }
                }

                LoginLinkedin.LoginByRandomUser(this.RunPage, this.LinkedinLoginPageUrl, this.LinkedinLoginSucceedCheckUrl);

                List<string> allPersonPageUrls = null;
                localLogFileName = "_" + this.GetKeyWords(seedRow) + "_personPageUrl";
                if (SysConfig.SysExecuteType == SysExecuteType.Produce || isNewDo)
                {
                    //如果是生产环境，那么直接爬取个人详情页
                    allPersonPageUrls = ProcessPersonPage.GetAllPersonPageUrls(this.RunPage, allPersonPageUrlInfos, this.GetLoginName(seedRow), this.GetLoginPassword(seedRow));
                    this.RunPage.SaveInfoToMiddleFile(localLogFileName, "personUrl", allPersonPageUrls);
                    isNewDo = true;
                }
                else
                {
                    //读取历史生成的个人网页网址
                    allPersonPageUrls = this.RunPage.TryGetInfoFromMiddleFile(localLogFileName, "personUrl");
                    if (allPersonPageUrls == null)
                    {
                        allPersonPageUrls = ProcessPersonPage.GetAllPersonPageUrls(this.RunPage, allPersonPageUrlInfos, this.GetLoginName(seedRow), this.GetLoginPassword(seedRow));
                        this.RunPage.SaveInfoToMiddleFile(localLogFileName, "personUrl", allPersonPageUrls);
                        isNewDo = true;
                    }
                }

                List<Dictionary<string, string>> personInfoList = ProcessPersonPage.GetPersonInfoFromLocalPages(this.RunPage, allPersonPageUrls, true, this.GetKeyWords(seedRow));
                 
                string personInfosResultFilePath = this.RunPage.GetFilePath("_SearchResult_Baidu2Linkedin_" + this.GetKeyWords(seedRow) + ".xlsx", this.RunPage.GetExportDir());

                ProcessPersonPage.SavePersonInfoToFile(this.RunPage, personInfoList, personInfosResultFilePath);
            }
            catch (Exception ex)
            {
                throw ex;
            }
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
                string pageUrl = seedPageUrl + "&pn=" + pageIndex * 10;
                this.GetCurrentPageAndNextPageUrl(seedPageUrl, keyWords, pageUrl, allListPageUrls);
                hasNextPage = this.HasNextPage(pageUrl);
                pageIndex++;
            }
        }

        private bool HasNextPage(string lastPageUrl)
        {
            string pageDir = this.RunPage.GetDetailSourceFileDir();
            string localFilePath = this.RunPage.GetFilePath(lastPageUrl, pageDir);
            StreamReader tr = new StreamReader(localFilePath, Encoding.UTF8);
            string webPageHtml = tr.ReadToEnd();

            HtmlAgilityPack.HtmlDocument pageHtmlDoc = new HtmlAgilityPack.HtmlDocument();
            pageHtmlDoc.LoadHtml(webPageHtml);

            HtmlNodeCollection nPageNodes = pageHtmlDoc.DocumentNode.SelectNodes("//a[@class=\"n\"]");
            if (nPageNodes == null || nPageNodes.Count == 0)
            {
                return false;
            }
            else
            {
                foreach (HtmlNode nPageNode in nPageNodes)
                {
                    if (nPageNode.InnerText.Contains("下一页"))
                    {
                        return true;
                    }
                }
                return false;
            }
        }

        /// <summary>
        /// 获取当前列表页及下一页地址
        /// </summary>
        /// <param name="listPageUrl"></param>
        /// <returns></returns>
        private void GetCurrentPageAndNextPageUrl(string seedPageUrl, string keyWords, string listPageUrl, List<string> allListPageUrls)
        {
            string localFilePath = this.RunPage.GetFilePath(listPageUrl, this.RunPage.GetDetailSourceFileDir());

            if (!File.Exists(localFilePath))
            {
                this.GetPageFromBaidu(listPageUrl, keyWords); 
                allListPageUrls.Add(listPageUrl);
            }
            else
            {
                allListPageUrls.Add(listPageUrl);
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

                HtmlNodeCollection divNodes = pageHtmlDoc.DocumentNode.SelectNodes("//div[@class=\"result c-container \"]");
                if (divNodes == null || divNodes.Count == 0)
                {
                    this.RunPage.InvokeAppendLogText("没有在Baidu中搜索到匹配项, keyWords = " + keyWords, LogLevelType.Normal, true);
                }
                else
                {
                    HtmlNodeCollection pageNodes = pageHtmlDoc.DocumentNode.SelectNodes("//div[@id=\"page\"]/a");

                    if (pageNodes == null || pageNodes.Count == 0)
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
                throw new Exception("抓取Baidu列表页失败", ex);
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
            return this.BaiduUrlPrefix + firstPageUrl.ToString();
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
        private List<Dictionary<string, string>> GetPersonPageUrlsFromListPages(string localDir, List<string> listPageUrls,string keyWords)
        {
            try
            {
                List<Dictionary<string, string>> allPersonPageUrlInfos = new List<Dictionary<string, string>>();
                foreach (string listPageUrl in listPageUrls)
                {
                    this.GetPersonPageUrls(localDir, listPageUrl, allPersonPageUrlInfos, keyWords);
                }
                return allPersonPageUrlInfos;
            }
            catch (Exception ex)
            {
                this.RunPage.InvokeAppendLogText(ex.Message, LogLevelType.Error, true);
                throw new Exception("解析列表页获取个人页面地址失败", ex);
            }
        }

        private void GetPersonPageUrls(string localDir, string listPageUrl, List<Dictionary<string, string>> allPersonPageUrlInfos,string keyWords)
        {
            try
            {
                string listPageLocalPath = this.RunPage.GetFilePath(listPageUrl, localDir);
                HtmlAgilityPack.HtmlDocument pageHtmlDoc = HtmlDocumentHelper.Load(listPageLocalPath);
                HtmlNodeCollection allDivNodes = pageHtmlDoc.DocumentNode.SelectNodes("//div[@class=\"result c-container \"]");
                foreach (HtmlNode divNode in allDivNodes)
                {
                    string linkedinUrlPart = HtmlDocumentHelper.TryGetNodeInnerText(divNode, "./div[@class=\"f13\"]/a", true, true, null, null);
                    if (linkedinUrlPart == null)
                    {
                        linkedinUrlPart = HtmlDocumentHelper.TryGetNodeInnerText(divNode, "./div/div[@class=\"f13\"]/a", true, true, null, null);
                    }

                    string abstractText  = HtmlDocumentHelper.TryGetNodeInnerText(divNode, true, true, null, null); 

                    if (linkedinUrlPart != null && linkedinUrlPart.Contains(".linkedin.com/in/") && abstractText != null && abstractText.ToLower().Contains(keyWords.ToLower()))
                    {
                        try
                        {
                            string personName = HtmlDocumentHelper.TryGetNodeInnerText(divNode, "./h3/a", true, true, null, null);
                            string personUrl = HtmlDocumentHelper.TryGetNodeAttributeValue(divNode, "./h3/a", "href", true, true, null, null);
                            foreach (string postfix in this.BaiduLinkedinItemPostfix)
                            {
                                personName = personName.Replace(postfix, "").Trim();
                            }
                            Dictionary<string, string> personPageUrlInfo = new Dictionary<string, string>();
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
                throw new Exception("解析Baidu列表页出错, listPageUrl = +" + listPageUrl, ex);
            }
        }
    }
}