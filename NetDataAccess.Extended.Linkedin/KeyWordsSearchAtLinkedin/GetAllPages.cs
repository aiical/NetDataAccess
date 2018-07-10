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

namespace NetDataAccess.Extended.Linkedin.KeyWordsSearchAtLinkedin
{
    /// <summary>
    /// GetAllListPage
    /// </summary>
    [System.Runtime.InteropServices.ComVisibleAttribute(true)]
    public class GetAllPages : ExternalRunWebPage
    { 
        private string _KeyWords = null;
        private string KeyWords
        {
            get
            {
                return _KeyWords;
            }
            set
            {
                _KeyWords = value;
            }
        }

        private string _LoginName = null;
        private string LoginName
        {
            get
            {
                return _LoginName;
            }
            set
            {
                _LoginName = value;
            }
        }

        private string _LoginPassword = null;
        private string LoginPassword
        {
            get
            {
                return _LoginPassword;
            }
            set
            {
                _LoginPassword = value;
            }
        }

        public string _SeedPageUrl = null;
        public string SeedPageUrl
        {
            get
            {
                return this._SeedPageUrl;
            }
            set
            {
                this._SeedPageUrl = value;
            }
        }

        public override bool BeforeAllGrab()
        {
            try
            {
                string[] ps = this.Parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                string loginPageUrl = ps[0];
                string loginSucceedCheckUrl = ps[1];

                this.RunPage.MustReGrab = true;

                if (SysConfig.SysExecuteType == SysExecuteType.Produce)
                {
                    //如果是生产环境，那么直接爬取列表页
                    LoginLinkedin.Login(this.RunPage, loginPageUrl, loginSucceedCheckUrl);
                }
                else
                {
                    //读取历史爬取的列表页地址文件 
                    if (this.RunPage.TryGetInfoFromMiddleFile("login", "login") == null)
                    {
                        LoginLinkedin.Login(this.RunPage, loginPageUrl, loginSucceedCheckUrl);
                        this.RunPage.SaveInfoToMiddleFile("login", "login", new List<string>());
                    }
                }
                this.RunPage.BeginGrab();

                return false;
            }
            catch (Exception ex)
            {
                throw new Exception("登录Linkedin失败! ", ex);
            }
        }

        public override bool AfterAllGrab(IListSheet listSheet)
        {
            this.GetSeedInfoFromListSheet(listSheet);

            //下一步必须执行
            bool isNewDo = false;
            string localLogFileName = null;

            List<string> allListPageUrls = null;
            localLogFileName = this.LoginName + "_" + this.KeyWords + "_listPageUrl";
            if (SysConfig.SysExecuteType == SysExecuteType.Produce)
            {
                //如果是生产环境，那么直接爬取列表页
                allListPageUrls = this.GetAllListPages(this.Parameters, this.SeedPageUrl);
                this.RunPage.SaveInfoToMiddleFile(localLogFileName, "listPageUrl", allListPageUrls);
                isNewDo = true;
            }
            else
            {
                //读取历史爬取的列表页地址文件 
                allListPageUrls = this.RunPage.TryGetInfoFromMiddleFile(localLogFileName, "listPageUrl");
                if (allListPageUrls == null)
                {
                    allListPageUrls = this.GetAllListPages(this.Parameters, this.SeedPageUrl);
                    this.RunPage.SaveInfoToMiddleFile(localLogFileName, "listPageUrl", allListPageUrls);
                    isNewDo = true;
                }

            }

            List<Dictionary<string, string>> allPersonPageUrlInfos = null;
            localLogFileName = this.LoginName + "_" + this.KeyWords + "_personPageUrlInfo";
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

            List<string> allPersonPageUrls = null;
            localLogFileName = this.LoginName + "_" + this.KeyWords + "_personPageUrl";
            if (SysConfig.SysExecuteType == SysExecuteType.Produce || isNewDo)
            {
                //如果是生产环境，那么直接爬取个人详情页
                allPersonPageUrls = ProcessPersonPage.GetAllPersonPageUrls(this.RunPage, allPersonPageUrlInfos, this.LoginName, this.LoginPassword);
                this.RunPage.SaveInfoToMiddleFile(localLogFileName, "personUrl", allPersonPageUrls);
                isNewDo = true;
            }
            else
            {
                //读取历史生成的个人网页网址
                allPersonPageUrls = this.RunPage.TryGetInfoFromMiddleFile(this.LoginName + "." + this.KeyWords + ".personPageUrl", "personUrl");
                if (allPersonPageUrls == null)
                {
                    allPersonPageUrls = ProcessPersonPage.GetAllPersonPageUrls(this.RunPage, allPersonPageUrlInfos, this.LoginName, this.LoginPassword);
                    this.RunPage.SaveInfoToMiddleFile(localLogFileName, "personUrl", allPersonPageUrls);
                    isNewDo = true;
                }
            }

            List<Dictionary<string, string>> personInfoList = ProcessPersonPage.GetPersonInfoFromLocalPages(this.RunPage, allPersonPageUrls, false, null);
             
            string personInfosFilePath = this.RunPage.GetFilePath("SearchResult_Linkedin2Linkedin_" + this.LoginName + "_" + this.KeyWords + ".xlsx", this.RunPage.GetExportDir());
            ProcessPersonPage.SavePersonInfoToFile(this.RunPage, personInfoList, personInfosFilePath);

            return true;
        }

        /// <summary>
        /// 从listSheet中加载获取爬取需要的输入参数
        /// </summary>
        /// <param name="listSheet"></param>
        private void GetSeedInfoFromListSheet(IListSheet listSheet)
        {
            Dictionary<string, string> listRow = listSheet.GetRow(0);
            this.KeyWords = listRow["keyWords"];
            this.LoginName = listRow["loginName"];
            this.LoginPassword = listRow["loginPassword"];
            this.SeedPageUrl = listRow[SysConfig.DetailPageUrlFieldName];
        }

        /// <summary>
        /// 循环获取所有列表页
        /// </summary>
        /// <param name="nextPageUrl"></param>
        private void GetListPageOneByOne(string nextPageUrl, List<string> allListPageUrls)
        {
            while (nextPageUrl != null)
            {
                nextPageUrl = this.GetCurrentPageAndNextPageUrl(nextPageUrl, allListPageUrls);
                Thread.Sleep(3000);
            }
        }

        /// <summary>
        /// 获取当前列表页及下一页地址
        /// </summary>
        /// <param name="listPageUrl"></param>
        /// <returns></returns>
        private string GetCurrentPageAndNextPageUrl(string listPageUrl, List<string> allListPageUrls)
        {
            string tabName = "ListPage";
            WebBrowser webBrowser = this.RunPage.ShowWebPage(listPageUrl, tabName, SysConfig.WebPageRequestTimeout, false);
            string currentPageUrl = this.RunPage.InvokeGetWebBrowserPageUrl(webBrowser);
            string listPageHtml = this.RunPage.InvokeGetPageHtml(tabName);
            string localFilePath = this.RunPage.GetFilePath(currentPageUrl, this.RunPage.GetDetailSourceFileDir());
            this.RunPage.SaveFile(listPageHtml, localFilePath, Encoding.UTF8);
            allListPageUrls.Add(currentPageUrl);

            string scriptMethodCode = "function myGetNextPageUrl(){"
                + "var nextLi = $('li.next');"
                + "if(nextLi.length == 0){"
                + "return '';"
                + "}"
                + "else{"
                + "return $(nextLi[0]).children('a').attr('href');"
                + "}"
                + "}";

            this.RunPage.InvokeAddScriptMethod(webBrowser, scriptMethodCode, this);
            string nextPageUrl = CommonUtil.UrlDecodeSymbolAnd((string)this.RunPage.InvokeDoScriptMethod(webBrowser, "myGetNextPageUrl", null));
            if (nextPageUrl != null && nextPageUrl.Length > 0)
            {
                return nextPageUrl;
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// 获取所有列表页内容
        /// </summary>
        /// <param name="parameters"></param>
        /// <param name="listSheet"></param>
        private List<string> GetAllListPages(string parameters, string seedPageUrl)
        {
            try
            {
                List<string> allListPageUrls = new List<string>();
                string pageDir = this.RunPage.GetDetailSourceFileDir();
                string localFilePath = this.RunPage.GetFilePath(seedPageUrl, pageDir);
                StreamReader tr = new StreamReader(localFilePath, Encoding.UTF8);
                string webPageHtml = tr.ReadToEnd();

                string limitAlert = "商业用途搜索次数已达到上限";

                if (webPageHtml.Contains(limitAlert))
                {
                    throw new Exception(limitAlert);
                }

                HtmlAgilityPack.HtmlDocument pageHtmlDoc = new HtmlAgilityPack.HtmlDocument();
                pageHtmlDoc.LoadHtml(webPageHtml);


                HtmlNodeCollection liNodes = pageHtmlDoc.DocumentNode.SelectNodes("//ul[@class=\"pagination\"]/li");
                if (liNodes.Count == 0)
                {
                    throw new Exception("没有找到符合条件的人");
                }
                else
                {
                    if (liNodes.Count == 1)
                    {
                        allListPageUrls.Add(seedPageUrl);
                    }
                    else
                    {
                        string firstPageUrl = this.GetFirstListPageUrl(liNodes);
                        if (firstPageUrl != null)
                        {
                            this.GetListPageOneByOne(firstPageUrl, allListPageUrls);
                        }
                    }
                }
                return allListPageUrls;
            }
            catch (Exception ex)
            {
                this.RunPage.InvokeAppendLogText(ex.Message, LogLevelType.Error, true);
                throw new Exception("抓取列表页失败", ex);
            }
        }

        /// <summary>
        /// 获取列表页首页地址
        /// </summary>
        /// <param name="liNodes"></param>
        /// <returns></returns>
        private string GetFirstListPageUrl(HtmlNodeCollection liNodes)
        {
            foreach (HtmlNode liNode in liNodes)
            {
                if (liNode.GetAttributeValue("class", "") != "active")
                {
                    string fullUrl = CommonUtil.UrlDecodeSymbolAnd(liNode.SelectSingleNode("./a").GetAttributeValue("href", ""));
                    string[] urlSplits = fullUrl.Split(new string[] { "?" }, StringSplitOptions.RemoveEmptyEntries);
                    string[] parameters = urlSplits[1].Split(new string[] { "&amp;" }, StringSplitOptions.RemoveEmptyEntries);
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
                        if (p.StartsWith("page_num="))
                        {
                            firstPageUrl.Append("page_num=1");
                        }
                        else
                        {
                            firstPageUrl.Append(p);
                        }
                    }
                    return firstPageUrl.ToString();
                }
            }
            return null;
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
                if (p.StartsWith("page_num="))
                {
                    return p.Replace("page_num=", "");
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
                HtmlNodeCollection allLiNodes = pageHtmlDoc.DocumentNode.SelectNodes("//ol[@class=\"search-results\"]/li");
                foreach (HtmlNode liNode in allLiNodes)
                {
                    if (liNode.GetAttributeValue("class", "").Contains("people"))
                    {
                        HtmlNode personLinkNode = liNode.SelectSingleNode("./div[@class=\"bd\"]/h3/a");
                        string personUrl = CommonUtil.UrlDecodeSymbolAnd(personLinkNode.GetAttributeValue("href", ""));
                        string personName = personLinkNode.InnerText.Trim();
                        Dictionary<string, string> personPageUrlInfo = new Dictionary<string, string>();
                        personPageUrlInfo.Add("personUrl", personUrl);
                        personPageUrlInfo.Add("personName", personName);
                        allPersonPageUrlInfos.Add(personPageUrlInfo);
                    }
                }
            }
            catch (Exception ex)
            {
                this.RunPage.InvokeAppendLogText(ex.Message, LogLevelType.Error, true);
                throw new Exception("解析列表页出错, listPageUrl = +" + listPageUrl, ex);
            }
        }
    }
}