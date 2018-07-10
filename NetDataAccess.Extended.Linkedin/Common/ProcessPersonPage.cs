using HtmlAgilityPack;
using mshtml;
using NetDataAccess.Base.Common;
using NetDataAccess.Base.Config;
using NetDataAccess.Base.EnumTypes;
using NetDataAccess.Base.UI;
using NetDataAccess.Base.Writer;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace NetDataAccess.Extended.Linkedin.Common
{
    public class ProcessPersonPage
    { 
        #region 根据网址和个人姓名，从linkedin网站获取个人信息网页，批量处理
        public static List<string> GetAllPersonPageUrls(IRunWebPage runPage, List<Dictionary<string, string>> allPersonPageUrlInfos, string loginName, string loginPassword)
        {
            List<string> allPersonPageUrls = new List<string>();
            foreach (Dictionary<string, string> personPageUrlInfo in allPersonPageUrlInfos)
            {
                try
                {
                    string personUrl = personPageUrlInfo["personUrl"];
                    string personName = personPageUrlInfo["personName"];
                    GetPersonPage(runPage, personUrl, personName, loginName, loginPassword, false, 0); 
                    allPersonPageUrls.Add(personUrl);
                }
                catch (Exception ex)
                {
                    runPage.InvokeAppendLogText(ex.Message, LogLevelType.Error, true);
                    //继续爬取
                }
            }
            return allPersonPageUrls;
        }
        #endregion

        #region 根据网址和个人姓名，从linkedin网站获取个人信息网页，批量处理
        public static List<Dictionary<string, string>> GetAllPersonPages(IRunWebPage runPage, List<Dictionary<string, string>> allPersonPageUrlInfos)
        {
            return GetAllPersonPages(runPage, allPersonPageUrlInfos, null, null);
        }
        #endregion

        #region 根据网址和个人姓名，从linkedin网站获取个人信息网页，批量处理
        public static List<Dictionary<string, string>> GetAllPersonPages(IRunWebPage runPage, List<Dictionary<string, string>> allPersonPageUrlInfos, string loginName, string loginPassword)
        {
            List<Dictionary<string, string>> allPersonPageInfosWithJustDownloadMark = new List<Dictionary<string, string>>();
            foreach (Dictionary<string, string> personPageUrlInfo in allPersonPageUrlInfos)
            {
                try
                {
                    string personUrl = personPageUrlInfo["personUrl"];
                    string personName = personPageUrlInfo["personName"];
                    bool isJustDownload = GetPersonPage(runPage, personUrl, personName, loginName, loginPassword, false,0 );
                    Dictionary<string, string> personInfo = new Dictionary<string, string>();
                    personPageUrlInfo.Add("isJustDownload", isJustDownload.ToString());
                    allPersonPageInfosWithJustDownloadMark.Add(personPageUrlInfo);
                }
                catch (Exception ex)
                {
                    runPage.InvokeAppendLogText(ex.Message, LogLevelType.Error, true);
                    //继续爬取
                }
            }
            return allPersonPageInfosWithJustDownloadMark;
        }
        #endregion

        #region 根据当前人展开找到其他相关人员的Url
        public static void GetRelatedPersonInfos(IRunWebPage runPage, List<Dictionary<string, string>> personPageInfos, Dictionary<string, string> allRelatedPersonUrlInfos, List<Dictionary<string, string>> allCheckedRelatedPersonInfos, string keyWords)
        { 
            foreach (Dictionary<string, string> personPageInfo in personPageInfos)
            {
                try
                {
                    //被查看关联人员的页面是不是刚刚获取到的
                    string personUrl = personPageInfo["personUrl"];
                    bool isJustDownload = Boolean.Parse(personPageInfo["isJustDownload"]);
                    string personName = personPageInfo["personName"];

                    GetRelatedPersonInfos(runPage, personUrl, personName, isJustDownload, allRelatedPersonUrlInfos, allCheckedRelatedPersonInfos, keyWords,  0);
                }
                catch (Exception ex)
                {
                    throw new Exception("获取关联的人员信息出错, keyWords = " + keyWords, ex);
                }
            }
        }
        private static void GetRelatedPersonInfos(IRunWebPage runPage, string personPageInfoId, string personName, bool sourcePageIsJustDownload, Dictionary<string, string> allRelatedPersonUrlInfos, List<Dictionary<string, string>> allCheckedRelatedPersonInfos, string keyWords, int levelNum)
        {
            runPage.InvokeAppendLogText("查找'" + personName + "'的'看过本页的会员还看了'. (levelNum = " + levelNum.ToString() + ")", LogLevelType.System, true);
            if (!sourcePageIsJustDownload)
            {
                //那么重新获取一遍这个页面
                GetPersonPage(runPage, personPageInfoId, personName, null, null, true, 0);
            }

            string relatedPersonListFileName = "_RelatedPersonUrlInfos_" + personPageInfoId;
            List<Dictionary<string, string>> relatedPersonUrlInfos = runPage.TryGetInfoFromMiddleFile(relatedPersonListFileName, new string[] { "relatedPersonInfoId", "relatedPersonUrl", "relatedPersonName", "isJustDownload" });
            if (!sourcePageIsJustDownload || relatedPersonUrlInfos == null)
            {
                try
                {
                    relatedPersonUrlInfos = GetRelatedPersonUrlInfosFromPage(runPage, personPageInfoId);
                }
                catch (Exception ex)
                {
                    throw new Exception("从本地的个人页面中获取此人的关联的联系人出错, keyWords = " + keyWords + ", personPageInfoId = " + personPageInfoId, ex);
                }
                foreach (Dictionary<string, string> relatedPersonUrlInfo in relatedPersonUrlInfos)
                {
                    string relatedPersonInfoId = relatedPersonUrlInfo["relatedPersonInfoId"];
                    string relatedPersonUrl = relatedPersonUrlInfo["relatedPersonUrl"];
                    string relatedPersonName = relatedPersonUrlInfo["relatedPersonName"];
                    try
                    {
                        bool isRelatedPageJustDownload = GetRelatedPersonPage(runPage, relatedPersonInfoId, relatedPersonUrl, relatedPersonName, null, null);
                        relatedPersonUrlInfo["isJustDownload"] = isRelatedPageJustDownload.ToString();
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("获取个人页面出错, keyWords = " + keyWords + ", relatedPersonUrl = " + relatedPersonUrl + ", relatedPersonName = " + relatedPersonName, ex);
                    }
                }

                //保存爬取下来此用户对应的关联用户的网址信息
                runPage.SaveInfoToMiddleFile(relatedPersonListFileName, new string[] { "relatedPersonInfoId", "relatedPersonUrl", "relatedPersonName", "isJustDownload" }, relatedPersonUrlInfos);
            }

            //逐个页面分析，看看是否包含要查询的关键字，把包含了关键字的信息爬取下来（这里是不是需要限定一下，只在工作目前就职公司、曾经就职公司等中搜索关键字）
            string checkedRelatedPersonListFileName = "_CheckedRelatedPersonUrlInfos_" + personPageInfoId + "_" + keyWords;
            List<Dictionary<string, string>> checkedRelatedPersonInfos = runPage.TryGetInfoFromMiddleFile(checkedRelatedPersonListFileName, new string[] { "relatedPersonInfoId", "relatedPersonUrl", "relatedPersonName", "isJustDownload" });
            if (!sourcePageIsJustDownload || checkedRelatedPersonInfos == null)
            {
                try
                {
                    checkedRelatedPersonInfos = CheckRelatedPersonInfos(runPage, relatedPersonUrlInfos, allRelatedPersonUrlInfos, keyWords);
                }
                catch (Exception ex)
                {
                    throw new Exception("判断是否是符合keyWords条件的人员, keyWords = " + keyWords + ", personPageInfoId = " + personPageInfoId, ex);
                }

                //从爬取下来的关联用户信息里，找到符合关键字的用户信息记录下来（按照某个命名规则记录在文件里，例如personPageUrl+keyWords），并递归找相关人
                runPage.SaveInfoToMiddleFile(checkedRelatedPersonListFileName, new string[] { "relatedPersonInfoId", "relatedPersonUrl", "relatedPersonName", "isJustDownload" }, checkedRelatedPersonInfos);
            }

            //记录下本次检查过的网页
            foreach (Dictionary<string, string> relatedPersonUrlInfo in relatedPersonUrlInfos)
            {
                string relatedPersonInfoId = relatedPersonUrlInfo["relatedPersonInfoId"];
                if (!allRelatedPersonUrlInfos.ContainsKey(relatedPersonInfoId))
                {
                    string relatedPersonUrl = relatedPersonUrlInfo["relatedPersonUrl"];
                    allRelatedPersonUrlInfos.Add(relatedPersonInfoId, relatedPersonUrl);
                }
            }

            //记录下符合条件的网页
            foreach (Dictionary<string, string> checkedRelatedPersonInfo in checkedRelatedPersonInfos)
            {
                string checkedRelatedPersonInfoId = checkedRelatedPersonInfo["relatedPersonInfoId"];
                Dictionary<string, string> checkedPersonInfoInfo = new Dictionary<string, string>();
                checkedPersonInfoInfo.Add("checkedRelatedPersonInfoId", checkedRelatedPersonInfoId);
                checkedPersonInfoInfo.Add("personUrl", checkedPersonInfoInfo["relatedPersonUrl"]);
                checkedPersonInfoInfo.Add("personName", checkedPersonInfoInfo["relatedPersonName"]);
                checkedPersonInfoInfo.Add("levelNum", levelNum.ToString());
                allCheckedRelatedPersonInfos.Add(checkedPersonInfoInfo);
            }

            runPage.InvokeAppendLogText("找到 " + checkedRelatedPersonInfos.Count.ToString() + " 个符合条件的'" + personName + "'的'看过本页的会员还看了'", LogLevelType.System, true);

            //递归找到相关人的相关人
            levelNum++;
            foreach (Dictionary<string, string> checkedRelatedPersonInfo in checkedRelatedPersonInfos)
            {
                string relatedPersonInfoId = checkedRelatedPersonInfo["relatedPersonInfoId"];
                string relatedPersonUrl = checkedRelatedPersonInfo["relatedPersonUrl"];
                string relatedPersonName = checkedRelatedPersonInfo["relatedPersonName"];
                bool isRelatedPageJustDownload = Boolean.Parse(checkedRelatedPersonInfo["isJustDownload"]);
                try
                {
                    GetRelatedPersonInfos(runPage, relatedPersonInfoId, relatedPersonName, isRelatedPageJustDownload, allRelatedPersonUrlInfos, allCheckedRelatedPersonInfos, keyWords, levelNum);
                }
                catch (Exception ex)
                {
                    throw new Exception("递归获取关联人员出错, keyWords = " + keyWords + ", relatedPersonInfoId = " + relatedPersonInfoId, ex);
                }
            }
        }

        /// <summary>
        /// 判断此人是否匹配keyWords
        /// </summary>
        /// <param name="relatedPersonUrlInfos"></param>
        /// <returns></returns>
        private static List<Dictionary<string, string>> CheckRelatedPersonInfos(IRunWebPage runPage, List<Dictionary<string, string>> relatedPersonUrlInfos,Dictionary<string,string> allRelatedPersonUrlInfos,string keyWords)
        {
            List<Dictionary<string, string>> checkedRelatedPersonInfos = new List<Dictionary<string, string>>();
            foreach (Dictionary<string, string> relatedPersonUrlInfo in relatedPersonUrlInfos)
            {
                string relatedPersonInfoId = relatedPersonUrlInfo["relatedPersonInfoId"];
                if (!allRelatedPersonUrlInfos.ContainsKey(relatedPersonInfoId) )
                {
                    string relatedPersonUrl = CheckRelatedPersonInfo(runPage, relatedPersonInfoId, keyWords);
                    if (relatedPersonUrl != null)
                    {
                        relatedPersonUrlInfo["relatedPersonUrl"] = relatedPersonUrl;
                        checkedRelatedPersonInfos.Add(relatedPersonUrlInfo);
                    }
                }
            }
            return checkedRelatedPersonInfos;
        }
        private static string CheckRelatedPersonInfo(IRunWebPage runPage, string relatedPersonInfoId, string keyWords)
        {
            string fileDir = runPage.GetDetailSourceFileDir();
            string filePath = runPage.GetFilePath(relatedPersonInfoId, fileDir);
            HtmlAgilityPack.HtmlDocument htmlDoc = HtmlDocumentHelper.Load(filePath);

            HtmlNodeCollection sectionCollection = htmlDoc.DocumentNode.SelectNodes("//section");
            bool isNewVersion = sectionCollection != null && sectionCollection.Count > 0;

            return CheckRelatedPersonInfo(runPage, htmlDoc, keyWords, isNewVersion);
        }
        private static string CheckRelatedPersonInfo(IRunWebPage runPage, HtmlAgilityPack.HtmlDocument htmlDoc, string keyWords, bool isNewVersion)
        {
            HtmlNode documentNode = htmlDoc.DocumentNode;
            HtmlNodeCollection relatedPersonNodes = documentNode.SelectNodes("//ul[class=\"browse-map-list\"]/li");
            List<Dictionary<string, string>> relatedPersonUrlInfos = new List<Dictionary<string, string>>();

            if (isNewVersion)
            {
                //目前就职 
                if (HtmlDocumentHelper.CheckNodeContainsText(documentNode, "//div[contains(@class, \"pv-top-card-section__information\")]/div[contains(@class, \"pv-top-card-section__experience\")]/h3[contains(@class, \"pv-top-card-section__company\")]", keyWords, false) //目前就职 
                    || HtmlDocumentHelper.CheckNodeContainsText(documentNode, "//section[contains(@class, \"experience-section\")]/ul/li", keyWords, false)) //工作经历
                {

                    //站内网址 
                    string znwz = CommonUtil.HtmlDecode(HtmlDocumentHelper.TryGetNodeInnerText(documentNode, "//section[contains(@class,\"ci-vanity-url\")]/div", true, true, null, null));
                
                    return znwz;
                }
                else
                {
                    return null;
                }
            }
            else
            {
                //目前就职 
                if (HtmlDocumentHelper.CheckNodeContainsText(documentNode, "//tr[@id=\"overview-summary-current\"]/td/ol/li", keyWords, false) //目前就职 
                    || HtmlDocumentHelper.CheckNodeContainsText(documentNode, "//tr[@id=\"overview-summary-past\"]/td/ol/li", keyWords, false) //曾经就职 
                    || HtmlDocumentHelper.CheckNodeContainsText(documentNode, "//div[@id=\"background-experience\"]/div/div", keyWords, false)) //工作经历
                {

                    //站内网址 
                    string znwz = CommonUtil.UrlDecode(HtmlDocumentHelper.TryGetNodeAttributeValue(documentNode, "//a[@class=\"view-public-profile\"]", "href", true));

                    return znwz;
                }
                else
                {
                    return null;
                }
            }
        }

        private static List<Dictionary<string, string>> GetRelatedPersonUrlInfosFromPage(IRunWebPage runPage, string personPageInfoId)
        {
            string fileDir = runPage.GetDetailSourceFileDir();
            string filePath = runPage.GetFilePath(personPageInfoId, fileDir);
            HtmlAgilityPack.HtmlDocument htmlDoc = HtmlDocumentHelper.Load(filePath);
            HtmlNode documentNode = htmlDoc.DocumentNode;
            HtmlNodeCollection relatedPersonNodes = documentNode.SelectNodes("//ul[@class=\"browse-map-list\"]/li");
            List<Dictionary<string, string>> relatedPersonUrlInfos = new List<Dictionary<string, string>>();
            if (relatedPersonNodes != null)
            {
                foreach (HtmlNode relatedPersonNode in relatedPersonNodes)
                {
                    HtmlNode relatedPersonLinkNode = relatedPersonNode.SelectSingleNode("./h4/a");

                    if (relatedPersonLinkNode != null)
                    {
                        string relatedPersonName = HtmlDocumentHelper.TryGetNodeInnerText(relatedPersonLinkNode, true, true, null, null);
                        string relatedPersonInfoId = HtmlDocumentHelper.TryGetNodeInnerText(relatedPersonNode, true, true, null, null);
                        string relatedPersonUrl = CommonUtil.UrlDecode(CommonUtil.UrlDecodeSymbolAnd(HtmlDocumentHelper.TryGetNodeAttributeValue(relatedPersonLinkNode, "href", true, true, null, null)));

                        Dictionary<string, string> relatedPersonUrlInfo = new Dictionary<string, string>();
                        relatedPersonUrlInfo.Add("relatedPersonInfoId", relatedPersonInfoId);
                        relatedPersonUrlInfo.Add("relatedPersonUrl", relatedPersonUrl);
                        relatedPersonUrlInfo.Add("relatedPersonName", relatedPersonName);
                        relatedPersonUrlInfos.Add(relatedPersonUrlInfo);
                    }
                    else
                    {
                        runPage.InvokeAppendLogText("即将到达会员档案浏览次数上限, 不能查看'看过本页的会员还看了'", LogLevelType.Error, true);
                        return relatedPersonUrlInfos;
                    }
                }
            }
            return relatedPersonUrlInfos;

        }
        #endregion

        #region 判断页面是否为新版的个人页面
        private static bool CheckIsNewVersionPersonPage(IRunWebPage runPage, WebBrowser webBrowser)
        {
            return (bool)webBrowser.Invoke(new CheckIsNewVersionPersonPageJavaScriptDelegate(CheckIsNewVersionPersonPageJavaScriptMethod), new object[] { runPage, webBrowser });
        }
        private delegate bool CheckIsNewVersionPersonPageJavaScriptDelegate(IRunWebPage runPage, WebBrowser webBrowser);
        private static bool CheckIsNewVersionPersonPageJavaScriptMethod(IRunWebPage runPage, WebBrowser webBrowser)
        {
            string scriptMethodCode = "function checkIsNewVersion(){"
                + "if($ == null){"
                + "return false;"
                + "}"
                + "else {"
                + "var allSections = $('section');"
                + "return (allSections.length == 0) ? false : true;"
                + "}"
                + "}";
            runPage.InvokeAddScriptMethod(webBrowser, scriptMethodCode, null);
            return (bool)runPage.InvokeDoScriptMethod(webBrowser, "checkIsNewVersion", null);
        }
        #endregion

        #region 判断页面是否为新版的个人页面
        private static void ExpandAllInfoInPage(IRunWebPage runPage, WebBrowser webBrowser)
        {
            webBrowser.Invoke(new ExpandAllInfoInPageJavaScriptDelegate(ExpandAllInfoInPageJavaScriptMethod), new object[] { runPage, webBrowser });
        }
        private delegate void ExpandAllInfoInPageJavaScriptDelegate(IRunWebPage runPage, WebBrowser webBrowser);
        private static void ExpandAllInfoInPageJavaScriptMethod(IRunWebPage runPage, WebBrowser webBrowser)
        {
            string scriptMethodCode = "function expanAllInfo(){"
                //联系方式，个人信息
                + "setTimeout(function(){ $('button[class*=\"contact-see-more-less\"]').click();}, 200);"
                //其他职位 
                + "setTimeout(function(){ $('section[class*=\"experience-section\"]').children('div[class*=\"pv-profile-section__actions-inline\"]').children('button').click();}, 400);"
                //教育经历 
                + "setTimeout(function(){ $('section[class*=\"education-section\"]').children('div[class*=\"pv-profile-section__actions-inline\"]').children('button').click();}, 600);"
                //还擅长
                + "setTimeout(function(){ $('button[aria-controls=\"featured-skills-expanded\"]').click();}, 800);"
                //其他语言
                + "setTimeout(function(){ $('button[aria-controls=\"languages-accomplishment-list\"]').click();}, 1000);"
                //出版作品
                + "setTimeout(function(){ $('section[class*=\"publications\"]').children('div[class*=\"pv-accomplishments-block__content\"]').children('div[class*=\"pv-profile-section__actions-inline \"]').children('button').click();}, 1200);"
                //所做项目 
                + "setTimeout(function(){ $('button[aria-controls=\"projects-accomplishment-list\"]').click();}, 1400);"
                //资格认证
                + "setTimeout(function(){ $('section[class*=\"certifications\"]').children('div[class*=\"pv-accomplishments-block__content\"]').children('div[class*=\"pv-profile-section__actions-inline \"]').children('button').click();}, 1600);"
                //工作经历 查看说明  
                + "setTimeout(function(){ $('button:contains(\"查看说明\")').click(); }, 1800);"
                + "}";
            runPage.InvokeAddScriptMethod(webBrowser, scriptMethodCode, null);
            runPage.InvokeDoScriptMethod(webBrowser, "expanAllInfo", null);
        }
        #endregion

        #region 根据网址和个人姓名，从linkedin网站获取个人信息网页
        private static bool GetPersonPage(IRunWebPage runPage, string personUrl, string personName, string loginName, string loginPassword, bool mustDownloadNow, int retryNum)
        {
            try
            {
                string tabName = "PersonPage";
                string localPersonFilePath = runPage.GetFilePath(personUrl, runPage.GetDetailSourceFileDir());
                if (mustDownloadNow || !File.Exists(localPersonFilePath))
                {
                    WebBrowser webBrowser = runPage.ShowWebPage(personUrl, tabName, SysConfig.WebPageRequestTimeout, false);

                    if (loginName != null && loginName != null)
                    {
                        LoginLinkedin.DoLoginMethod(runPage, webBrowser, loginName, loginPassword);
                    }

                    ProcessWebBrowser.AutoScroll(runPage, webBrowser, 2000, 1000, 500, 1000);

                    bool openRight = runPage.CheckOpenRightPage(webBrowser, new string[] { personName, "发送 InMail", "未找到会员资料", "您和此领英用户没有共同联系人" }, new string[] { "忘记密码" }, SysConfig.WebPageRequestTimeout, false);
                    if (!openRight)
                    {
                        throw new Exception("没有打开正确的页面, 原计划打开的页面是" + personUrl);
                    }

                    bool isNewVersion = CheckIsNewVersionPersonPage(runPage, webBrowser);
                    if (isNewVersion)
                    {
                        ExpandAllInfoInPage(runPage, webBrowser);
                        Thread.Sleep(3000);
                    }

                    string personPageHtml = runPage.InvokeGetPageHtml(tabName);
                    runPage.SaveFile(personPageHtml, localPersonFilePath, Encoding.UTF8);

                    //表示这是新获取的网页
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                runPage.InvokeAppendLogText(ex.Message, LogLevelType.Error, true);
                if (retryNum < 3)
                {
                    retryNum++;
                    runPage.InvokeAppendLogText("获取个人页面出错, personUrl = " + personUrl + ", 将要重试（重试次数=" + retryNum + "）", LogLevelType.Error, true);
                    return ProcessPersonPage.GetPersonPage(runPage, personUrl, personName, loginName, loginPassword, mustDownloadNow, retryNum);
                }
                else
                {
                    throw new Exception("获取个人页面出错, personUrl = " + personUrl + ", personName = " + personName, ex);
                }
            }
            finally
            {
                runPage.CloseWebPage("PersonPage");
            }
        }
        #endregion

        #region 根据网址和个人姓名，从linkedin网站获取个人信息网页
        private static bool GetRelatedPersonPage(IRunWebPage runPage, string personInfoId, string personPageUrl, string personName, string loginName, string loginPassword)
        {
            try
            {
                string tabName = "RelatedPersonPage";
                //使用的个人信息作为唯一id标志
                string localPersonFilePath = runPage.GetFilePath(personInfoId, runPage.GetDetailSourceFileDir());
                if (!File.Exists(localPersonFilePath))
                {
                    WebBrowser webBrowser = runPage.ShowWebPage(personPageUrl, tabName, SysConfig.WebPageRequestTimeout, false);

                    LoginLinkedin.DoLoginMethod(runPage, webBrowser, loginName, loginPassword);

                    ProcessWebBrowser.AutoScroll(runPage, webBrowser, 5000, 2000, 100, 1000);

                    runPage.CheckWebBrowserContainsForComplete(webBrowser, new string[] { personName, "发送 InMail", "未找到会员资料", "您和此领英用户没有共同联系人" }, SysConfig.WebPageRequestTimeout, false);
                    string personPageHtml = runPage.InvokeGetPageHtml(tabName);
                    runPage.SaveFile(personPageHtml, localPersonFilePath, Encoding.UTF8);

                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                runPage.InvokeAppendLogText(ex.Message, LogLevelType.Error, true);
                throw new Exception("获取个人页面出错, personPageUrl = " + personPageUrl + ", personPageUrl = " + personPageUrl + ", personName = " + personName, ex);
            }
        }
        #endregion

        #region 将所有个人信息输出到本地文件
        public static void SavePersonInfoToFile(IRunWebPage runPage, List<Dictionary<string, string>> personInfoList, string ewFilePath)
        {
            Dictionary<string, int> columnNameToIndex = CommonUtil.InitStringIndexDic(new string[]{
                "姓名", 
                "英文名",  
                "目前就职",   
                "职位",  
                "所在地",  
                "曾经就职",  
                "教育经历",  
                "电话",  
                "邮箱", 
                "微信", 
                "微信二维码", 
                "网站", 
                "工作经历",  
                "资格认证",  
                "支持的组织机构",  
                "出版作品",  
                "语言能力",  
                "技能",  
                "还擅长",  
                "个人信息",  
                "联系详情",  
                "参与组织", 
                "所做项目",  
                "推荐信",  
                "联系人",  
                "关注内容",  
                "站内网址"});
            ExcelWriter ew = new ExcelWriter(ewFilePath, "List", columnNameToIndex);

            foreach (Dictionary<string, string> row in personInfoList)
            {
                ew.AddRow(row);
            }
            ew.SaveToDisk();
        }
        #endregion

        #region 从本地的网页文件中，获取个人信息，批量处理
        public static List<Dictionary<string, string>> GetPersonInfoFromLocalPages(IRunWebPage runPage, List<Dictionary<string, string>> allPersonPageInfos, bool needCheckKeyWords, string keyWords)
        {
            List<Dictionary<string, string>> personInfoList = new List<Dictionary<string, string>>();
            Dictionary<string, string> personUrlDic = new Dictionary<string, string>();
            foreach (Dictionary<string, string> personPageInfo in allPersonPageInfos)
            {
                string personPageUrl = personPageInfo["personUrl"];
                Dictionary<string, string> personInfo = GetPersonInfoFromLocalPage(runPage, personUrlDic, personPageUrl, needCheckKeyWords, keyWords);
                if (personInfo != null)
                {
                    personInfoList.Add(personInfo);
                }
            }
            return personInfoList;
        }
        #endregion

        #region 从本地的网页文件中，获取个人信息，批量处理
        public static List<Dictionary<string, string>> GetPersonInfoFromLocalPages(IRunWebPage runPage, List<string> allPersonPageUrls, bool needCheckKeyWords, string keyWords)
        {
            List<Dictionary<string, string>> personInfoList = new List<Dictionary<string, string>>();
            Dictionary<string, string> personUrlDic = new Dictionary<string, string>();
            foreach (string personPageUrl in allPersonPageUrls)
            {
                Dictionary<string, string> personInfo = GetPersonInfoFromLocalPage(runPage, personUrlDic, personPageUrl, needCheckKeyWords, keyWords);
                if (personInfo != null)
                {
                    personInfoList.Add(personInfo);
                }
            }
            return personInfoList;
        }
        #endregion

        #region 从本地的网页文件中，获取个人信息
        private static Dictionary<string, string> GetPersonInfoFromLocalPage(IRunWebPage runPage,Dictionary<string,string>personUrlDic, string personPageUrl, bool needCheckKeyWords, string keyWords)
        {
            string fileDir = runPage.GetDetailSourceFileDir();
            string filePath = runPage.GetFilePath(personPageUrl, fileDir);
            HtmlAgilityPack.HtmlDocument htmlDoc = HtmlDocumentHelper.Load(filePath);
            HtmlNode documentNode = htmlDoc.DocumentNode;

            HtmlNodeCollection sectionCollection = documentNode.SelectNodes("//section");
            bool isNewVersion = sectionCollection != null && sectionCollection.Count > 0;

            if (!needCheckKeyWords || CheckRelatedPersonInfo(runPage, htmlDoc, keyWords, isNewVersion) != null)
            {
                string xm = "";
                string englishName = "";
                string mqjz = "";
                string zw = "";
                string szd = "";
                string cjjz = "";
                string jyjl = "";
                string dh = "";
                string yx = "";
                string wx = "";
                string wxewm = "";
                string wz = "";
                string gzjl = "";
                string zgrz = "";
                string zcdzzjg = "";
                string cbzp = "";
                string yynl = "";
                string jn = "";
                string hsc = "";
                string grxx = "";
                string lxxq = "";
                string cyzz = "";
                string szxm = "";
                string tjx = "";
                string lxr = "";
                string gznr = "";
                string znwz = "";

                //站内网址
                if (isNewVersion)
                {
                    znwz = CommonUtil.HtmlDecode(HtmlDocumentHelper.TryGetNodeInnerText(documentNode, "//section[contains(@class,\"ci-vanity-url\")]/div", true, true, null, null));
                }
                else
                {
                    znwz = CommonUtil.UrlDecode(HtmlDocumentHelper.TryGetNodeAttributeValue(documentNode, "//a[@class=\"view-public-profile\"]", "href", true));
                }

                if (znwz == null)
                {
                    throw new Exception("无法在个人网页里获取个人网址, personPageUrl = " + personPageUrl);
                }


                if (!personUrlDic.ContainsKey(znwz))
                {
                    personUrlDic.Add(znwz, null);
                    if (isNewVersion)
                    {
                        #region 解析新版本页面里的个人信息
                        //姓名 
                        HtmlNode nameNode = documentNode.SelectSingleNode("//div[contains(@class, \"pv-top-card-section__information\")]/h1[contains(@class,\"pv-top-card-section__name\")]");
                        if (nameNode != null)
                        {
                            string fullXM = nameNode.InnerText.Trim();
                            int enIndex = fullXM.IndexOf("(");
                            if (enIndex >= 0)
                            {
                                xm = fullXM.Substring(0, enIndex);
                                englishName = fullXM.Substring(enIndex + 1, fullXM.Length - enIndex - 2);
                            }
                            else
                            {
                                xm = fullXM;
                            }
                        }
                        else
                        {
                            runPage.InvokeAppendLogText("爬取的Linkedin个人网页内容无效, 无法找到姓名. personPageUrl = " + personPageUrl, LogLevelType.System, true);
                        }


                        //目前就职
                        mqjz = HtmlDocumentHelper.TryGetNodeInnerText(documentNode, "//div[contains(@class, \"pv-top-card-section__information\")]/div[contains(@class, \"pv-top-card-section__experience\")]/h3[contains(@class, \"pv-top-card-section__company\")]", true);

                        //职位
                        zw = HtmlDocumentHelper.TryGetNodeInnerText(documentNode, "//div[contains(@class, \"pv-top-card-section__information\")]/h2[contains(@class, \"pv-top-card-section__headline\")]", true);

                        //所在地
                        szd = HtmlDocumentHelper.TryGetNodeInnerText(documentNode, "//div[contains(@class, \"pv-top-card-section__information\")]/h3[contains(@class, \"pv-top-card-section__location\")]", true);

                        //曾经就职,新版本页面无此项

                        //教育经历
                        HtmlNodeCollection jyjlNodes = documentNode.SelectNodes("//section[contains(@class, \"education-section\")]/ul/li");
                        if (jyjlNodes != null)
                        {
                            StringBuilder jyjlBuilder = new StringBuilder();
                            foreach (HtmlNode jyjlNode in jyjlNodes)
                            {
                                string jyjlOne = HtmlDocumentHelper.TryGetNodeInnerText(jyjlNode, true, true, null, null).Replace("\n", " ");
                                jyjlBuilder.Append(CommonUtil.StringArrayToString(jyjlOne.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries), " ")); 
                                jyjlBuilder.Append("\r\n");
                            }
                            jyjl = jyjlBuilder.ToString().Replace("<!---->", "");
                        }

                        //电话，新版本取不到
                        dh = HtmlDocumentHelper.TryGetNodeInnerText(documentNode, "//section[contains(@id, \"ci-phone\")]", true, true, null, null);

                        //邮箱
                        yx = HtmlDocumentHelper.TryGetNodeInnerText(documentNode, "//section[contains(@id, \"ci-email\")]", true, true, null, null);

                        //微信二维码，新版本取不到

                        //微信
                        wx = HtmlDocumentHelper.TryGetNodeInnerText(documentNode, "//button[contains(@class, \"add-wechat-friend-btn\")]/span[@class!=\"visually-hidden\"]", true);

                        //网站
                        wz = HtmlDocumentHelper.TryGetNodeInnerText(documentNode,"//section[contains(@id, \"ci-websites\")]", true, true, null, null);

                        //工作经历
                        HtmlNodeCollection gzjlNodes = documentNode.SelectNodes("//section[contains(@class, \"experience-section\")]/ul/li");
                        if (gzjlNodes != null)
                        {
                            StringBuilder gzjlBuilder = new StringBuilder();
                            foreach (HtmlNode gzjlNode in gzjlNodes)
                            {
                                string gzjlOne = HtmlDocumentHelper.TryGetNodeInnerText(gzjlNode, true, true, null, null).Replace("\n", " ");
                                gzjlBuilder.Append(CommonUtil.StringArrayToString(gzjlOne.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries), " "));
                                gzjlBuilder.Append("\r\n");
                            }
                            gzjl = gzjlBuilder.ToString().Replace("<!---->", "");
                        } 

                        //资格认证
                        HtmlNodeCollection zgrzNodes = documentNode.SelectNodes("//div[@id=\"certifications-accomplishment-list\"]/ul/li");
                        if (zgrzNodes != null)
                        {
                            StringBuilder zgrzBuilder = new StringBuilder();
                            foreach (HtmlNode zgrzNode in zgrzNodes)
                            {
                                string zgrzOne = HtmlDocumentHelper.TryGetNodeInnerText(zgrzNode, true, true, null, null).Replace("\n", " ");
                                zgrzBuilder.Append(CommonUtil.StringArrayToString(zgrzOne.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries), " "));
                                zgrzBuilder.Append("\r\n");
                            }
                            zgrz = zgrzBuilder.ToString().Replace("<!---->", "");
                        }

                        //支持的组织机构，新版本未找到

                        //出版作品
                        HtmlNodeCollection cbzpNodes = documentNode.SelectNodes("//div[@id=\"publications-accomplishment-list\"]/ul/li");
                        if (cbzpNodes != null)
                        {
                            StringBuilder cbzpBuilder = new StringBuilder();
                            foreach (HtmlNode cbzpNode in cbzpNodes)
                            {
                                string cbzpOne = HtmlDocumentHelper.TryGetNodeInnerText(cbzpNode, true, true, null, null).Replace("\n", " ");
                                cbzpBuilder.Append(CommonUtil.StringArrayToString(cbzpOne.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries), " "));  
                                cbzpBuilder.Append("\r\n");
                            }
                            cbzp = cbzpBuilder.ToString().Replace("<!---->", "");
                        }

                        //语言能力
                        HtmlNodeCollection yynlNodes = documentNode.SelectNodes("//div[@id=\"languages-accomplishment-list\"]/ul/li");
                        if (yynlNodes != null)
                        {
                            StringBuilder yynlBuilder = new StringBuilder();
                            foreach (HtmlNode yynlNode in yynlNodes)
                            {
                                string yynlOne = HtmlDocumentHelper.TryGetNodeInnerText(yynlNode, true, true, null, null).Replace("\n", " ");
                                yynlBuilder.Append(CommonUtil.StringArrayToString(yynlOne.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries), " "));  
                                yynlBuilder.Append("\r\n");
                            }
                            yynl = yynlBuilder.ToString().Replace("<!---->", "");
                        }

                        //技能
                        HtmlNodeCollection jnNodes = documentNode.SelectNodes("//section[contains(@class, \"pv-featured-skills-section\")]/ul[contains(@class,\"pv-featured-skills-list\")]/li");
                        if (jnNodes != null)
                        {
                            List<String> jnList = new List<string>();
                            foreach (HtmlNode jnNode in jnNodes)
                            {
                                string jnName = HtmlDocumentHelper.TryGetNodeInnerText(jnNode, "./div[1]/div[1]/a[1]/div[1]/span[contains(@class, \"pv-skill-entity__skill-name\")]", true, true, null, null);
                                if (jnName != null)
                                {
                                    string jnCount = HtmlDocumentHelper.TryGetNodeInnerText(jnNode, "./div[1]/div[1]/a[1]/div[1]/span[contains(@class, \"pv-skill-entity__endorsement-count\")]", true, true, null, null);
                                    jnList.Add(jnName + "(" + jnCount + ");");
                                }
                                else
                                {
                                    jnName = HtmlDocumentHelper.TryGetNodeInnerText(jnNode, "./div[1]/div[1]/div[1]/span[contains(@class, \"pv-skill-entity__skill-name\")]", true, true, null, null);
                                    if (jnName != null)
                                    {
                                        jnList.Add(jnName+";");
                                    }
                                }
                            }
                            jn = CommonUtil.StringArrayToString(jnList.ToArray(), "");
                        }

                        //还擅长
                        HtmlNodeCollection hscNodes = documentNode.SelectNodes("//div[@id=\"featured-skills-expanded\"]/ul[contains(@class,\"pv-featured-skills-list\")]/li");
                        if (hscNodes != null)
                        {
                            List<String> hscList = new List<string>();
                            foreach (HtmlNode hscNode in hscNodes)
                            {
                                string hscName = HtmlDocumentHelper.TryGetNodeInnerText(hscNode, "./div[1]/div[1]/a[1]/div[1]/span[contains(@class, \"pv-skill-entity__skill-name\")]", true, true, null, null);
                                if (hscName != null)
                                {
                                    string hscCount = HtmlDocumentHelper.TryGetNodeInnerText(hscNode, "./div[1]/div[1]/a[1]/div[1]/span[contains(@class, \"pv-skill-entity__endorsement-count\")]", true, true, null, null);
                                    hscList.Add(hscName + "(" + hscCount + ");");
                                }
                                else
                                {
                                    hscName = HtmlDocumentHelper.TryGetNodeInnerText(hscNode, "./div[1]/div[1]/div[1]/span[contains(@class, \"pv-skill-entity__skill-name\")]", true, true, null, null);
                                    if (hscName != null)
                                    {
                                        hscList.Add(hscName + ";");
                                    }
                                }
                            }
                            hsc = CommonUtil.StringArrayToString(hscList.ToArray(), "");
                        }

                        //个人信息，新版本中未找到

                        //联系详情，新版本中未找到

                        //参与组织，新版本中未找到

                        //所做项目
                        HtmlNodeCollection szxmNodes = documentNode.SelectNodes("//div[@id=\"projects-accomplishment-list\"]/ul/li");
                        if (szxmNodes != null)
                        {
                            StringBuilder szxmBuilder = new StringBuilder();
                            foreach (HtmlNode szxmNode in szxmNodes)
                            {
                                string szxmOne = HtmlDocumentHelper.TryGetNodeInnerText(szxmNode, true, true, null, null).Replace("\n", " ");
                                szxmBuilder.Append(CommonUtil.StringArrayToString(szxmOne.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries), " ")); 
                                szxmBuilder.Append("\r\n");
                            }
                            szxm = szxmBuilder.ToString().Replace("<!---->", "");
                        } 

                        //推荐信
                        //直接全部文本摘录下来吗?????????????????????????这里只是爬取了推荐信的个数
                        tjx = HtmlDocumentHelper.TryGetNodeInnerText(documentNode, "//div[contains(@class, \"recommendations-inlining\")]/div[1]/ul", true, true, null, null);
                        tjx = tjx == null ? null : tjx.Replace("\n", "").Replace(" ", "").Replace("<!---->", "");
                      
                        //联系人, 这里只是爬取此用户有几个联系人
                        lxr = HtmlDocumentHelper.TryGetNodeInnerText(documentNode, "//h3[contains(@class, \"pv-top-card-section__connections\")]/span[1]", true, true, null, null);

                        //关注内容
                        HtmlNodeCollection gznrNodes = documentNode.SelectNodes("//section[contains(@class, \"interests-section\")]/ul/li");
                        if (gznrNodes != null)
                        {
                            StringBuilder gznrBuilder = new StringBuilder();
                            foreach (HtmlNode szxmNode in gznrNodes)
                            {
                                string gznrOne = HtmlDocumentHelper.TryGetNodeInnerText(szxmNode, true, true, null, null).Replace("\n", " ");
                                gznrBuilder.Append(CommonUtil.StringArrayToString(gznrOne.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries), " "));
                                gznrBuilder.Append("\r\n");
                            }
                            gznr = gznrBuilder.ToString().Replace("<!---->", "");
                        } 
                        #endregion
                    }
                    else
                    {
                        #region 解析旧版本页面里的个人信息
                        //姓名 
                        HtmlNode nameNode = documentNode.SelectSingleNode("//span[@class=\"full-name\"]");
                        if (nameNode != null)
                        {
                            xm = nameNode.FirstChild.InnerText.Trim();

                            //英文名 
                            HtmlNode englishNameNode = nameNode.SelectSingleNode("./span[@class=\"english-name\"]");
                            if (englishNameNode != null)
                            {
                                englishName = englishNameNode.InnerText.Replace("(", "").Replace(")", "").Trim();
                            }
                        }
                        else
                        {
                            runPage.InvokeAppendLogText("爬取的Linkedin个人网页内容无效, 无法找到姓名. personPageUrl = " + personPageUrl, LogLevelType.System, true);
                        }


                        //目前就职
                        mqjz = HtmlDocumentHelper.TryGetNodeInnerText(documentNode, "//tr[@id=\"overview-summary-current\"]/td/ol/li", true);

                        //职位
                        zw = HtmlDocumentHelper.TryGetNodeInnerText(documentNode, "//div[@id=\"headline\"]", true);

                        //所在地
                        szd = HtmlDocumentHelper.TryGetNodeInnerText(documentNode, "//div[@id=\"location\"]/dl/dd/span[@class=\"locality\"]", true);

                        //曾经就职
                        cjjz = HtmlDocumentHelper.JointNodeInnerText(documentNode, "//tr[@id=\"overview-summary-past\"]/td/ol/li", ";", true, true, null, null);

                        //教育经历
                        HtmlNodeCollection jyjlNodes = documentNode.SelectNodes("//div[@id=\"background-education\"]/div/div/div");
                        if (jyjlNodes != null)
                        {
                            StringBuilder jyjlBuilder = new StringBuilder();
                            foreach (HtmlNode jyjlNode in jyjlNodes)
                            {
                                jyjlBuilder.Append(HtmlDocumentHelper.TryGetNodeInnerText(jyjlNode, "./header/h4", true, true, "学校:", "; "));

                                jyjlBuilder.Append(HtmlDocumentHelper.JointNodeInnerText(jyjlNode, "./header/h5/span", " ", true, true, "学校:", "; "));

                                jyjlBuilder.Append(HtmlDocumentHelper.JointNodeInnerText(jyjlNode, "./span[@class=\"education-date\"]/time", "", true, true, "时间:", "; "));

                                jyjlBuilder.Append("\r\n");
                            }
                            jyjl = jyjlBuilder.ToString();
                        }

                        //电话
                        dh = HtmlDocumentHelper.JointNodeInnerText(documentNode, "//div[@id=\"phone\"]/div[@id=\"phone-view\"]/ul/li", "; ", true, true, null, null);

                        //邮箱
                        yx = HtmlDocumentHelper.JointNodeInnerText(documentNode, "//div[@id=\"email\"]/div[@id=\"email-view\"]/ul/li", "; ", true, true, null, null);

                        //微信二维码
                        wxewm = CommonUtil.UrlDecodeSymbolAnd(HtmlDocumentHelper.TryGetNodeAttributeValue(documentNode, "//div[@id=\"wechat\"]", "data-qr2-url", true));
                        wxewm = wxewm == null || wxewm.Length == 0 ? null : wxewm;

                        //微信
                        wx = HtmlDocumentHelper.TryGetNodeInnerText(documentNode, "//div[@id=\"wechat\"]", true);

                        //网站
                        HtmlNodeCollection websiteNodes = documentNode.SelectNodes("//div[@id=\"website\"]/div[@id=\"website-view\"]/ul/li/a");
                        if (websiteNodes != null)
                        {
                            List<String> wzList = new List<string>();
                            foreach (HtmlNode webSiteNode in websiteNodes)
                            {
                                string href = CommonUtil.UrlDecodeSymbolAnd(webSiteNode.GetAttributeValue("href", "").Trim());
                                string wzName = CommonUtil.HtmlDecode(webSiteNode.InnerText.Trim());
                                wzList.Add(wzName + ":" + href);
                            }
                            wz = CommonUtil.StringArrayToString(wzList.ToArray(), "\r\n");
                        }

                        //工作经历
                        HtmlNodeCollection gzjlNodes = documentNode.SelectNodes("//div[@id=\"background-experience\"]/div/div");
                        if (gzjlNodes != null)
                        {
                            StringBuilder gzjlBuilder = new StringBuilder();
                            foreach (HtmlNode gzjlNode in gzjlNodes)
                            {
                                gzjlBuilder.Append(HtmlDocumentHelper.TryGetNodeInnerText(gzjlNode, "./header/h4", true, true, "岗位:", "; "));

                                gzjlBuilder.Append("\r\n");
                                gzjlBuilder.Append(HtmlDocumentHelper.JointNodeInnerText(gzjlNode, "./header/h5", " ", true, true, "公司:", "; "));

                                gzjlBuilder.Append("\r\n");
                                gzjlBuilder.Append(HtmlDocumentHelper.JointNodeInnerText(gzjlNode, "./span[@class=\"experience-date-locale\"]/time", "", true, true, "时间:", "; "));

                                gzjlBuilder.Append("\r\n");
                                gzjlBuilder.Append(HtmlDocumentHelper.TryGetNodeInnerText(gzjlNode, "./span[@class=\"experience-date-locale\"]/span[@class=\"locality\"]", true, true, "地址:", "; "));

                                gzjlBuilder.Append("\r\n");
                                gzjlBuilder.Append(HtmlDocumentHelper.TryGetNodeInnerText(gzjlNode, "./p[@class=\"description summary-field-show-more\"]", true, true, "描述:", "; "));

                                gzjlBuilder.Append("\r\n");
                                gzjlBuilder.Append("\r\n");
                            }
                            gzjl = gzjlBuilder.ToString();
                        }

                        //资格认证
                        HtmlNodeCollection zgrzNodes = documentNode.SelectNodes("//div[@id=\"background-certifications\"]/div/div");
                        if (zgrzNodes != null)
                        {
                            StringBuilder zgrzBuilder = new StringBuilder();
                            foreach (HtmlNode zgrzNode in zgrzNodes)
                            {
                                zgrzBuilder.Append(HtmlDocumentHelper.TryGetNodeInnerText(zgrzNode, "./hgroup/h4", true, true, "", ", "));

                                zgrzBuilder.Append(HtmlDocumentHelper.JointNodeInnerText(zgrzNode, "./hgroup/h5", " ", true, true, "", ", "));

                                zgrzBuilder.Append(HtmlDocumentHelper.JointNodeInnerText(zgrzNode, "./span[@class=\"certification-date\"]/*", "", true, true, "", " "));

                                zgrzBuilder.Append("\r\n");
                            }
                            zgrz = zgrzBuilder.ToString();
                        }

                        //支持的组织机构
                        zcdzzjg = HtmlDocumentHelper.JointNodeInnerText(documentNode, "//div[@class=\"volunteering-organizations-view\"]/div/ul[@class=\"volunteering-listing\"]/li", "; ", true, true, null, null);

                        //出版作品
                        HtmlNodeCollection cbzpNodes = documentNode.SelectNodes("//div[@id=\"background-publications\"]/div/div");
                        if (cbzpNodes != null)
                        {
                            StringBuilder cbzpBuilder = new StringBuilder();
                            foreach (HtmlNode cbzpNode in cbzpNodes)
                            {
                                cbzpBuilder.Append(HtmlDocumentHelper.TryGetNodeInnerText(cbzpNode, "./hgroup/h4", true, true, "", ", "));

                                cbzpBuilder.Append(HtmlDocumentHelper.TryGetNodeInnerText(cbzpNode, "./hgroup/h5", true, true, "", ", "));

                                cbzpBuilder.Append(HtmlDocumentHelper.TryGetNodeInnerText(cbzpNode, "./span[@class=\"publication-date\"]", true, true, "", ""));

                                cbzpBuilder.Append("\r\n");
                            }
                            cbzp = cbzpBuilder.ToString();
                        }

                        //语言能力
                        HtmlNodeCollection yynlNodes = documentNode.SelectNodes("//div[@id=\"languages\"]/div[@id=\"languages-view\"]/ol/li");
                        if (yynlNodes != null)
                        {
                            List<String> yynlList = new List<string>();
                            foreach (HtmlNode yynlNode in yynlNodes)
                            {
                                string yy = HtmlDocumentHelper.TryGetNodeInnerText(yynlNode, "./h4", true);
                                string nl = HtmlDocumentHelper.TryGetNodeInnerText(yynlNode, "./div[@class=\"languages-proficiency\"]", true);
                                yynlList.Add(yy + ", " + nl + ";");
                            }
                            yynl = CommonUtil.StringArrayToString(yynlList.ToArray(), "\r\n");
                        }

                        //技能
                        HtmlNodeCollection jnNodes = documentNode.SelectNodes("//div[@id=\"profile-skills\"]/ul[@class=\"skills-section\"]/li/span[@class=\"skill-pill\"]");
                        if (jnNodes != null)
                        {
                            List<String> jnList = new List<string>();
                            foreach (HtmlNode jnNode in jnNodes)
                            {
                                string jnCount = HtmlDocumentHelper.TryGetNodeAttributeValue(jnNode, "./a[@class=\"endorse-count\"]/span[@class=\"num-endorsements\"]", "data-count", true);
                                string jnName = HtmlDocumentHelper.TryGetNodeInnerText(jnNode, "./span[@class=\"endorse-item-name\"]/span[@class=\"endorse-item-name-text\"]", true);
                                jnList.Add(jnName + "(" + jnCount + ");");
                            }
                            jn = CommonUtil.StringArrayToString(jnList.ToArray(), "");
                        }

                        //还擅长
                        HtmlNodeCollection hscNodes = documentNode.SelectNodes("//div[@id=\"profile-skills\"]/ul[@class=\"skills-section compact-view\"]/li/div/span[@class=\"skill-pill\"]");
                        if (hscNodes != null)
                        {
                            List<String> hscList = new List<string>();
                            foreach (HtmlNode hscNode in hscNodes)
                            {
                                string hscCount = HtmlDocumentHelper.TryGetNodeAttributeValue(hscNode, "./a[@class=\"endorse-count\"]/span[@class=\"num-endorsements\"]", "data-count", true);
                                string hscName = HtmlDocumentHelper.TryGetNodeInnerText(hscNode, "./span[@class=\"endorse-item-name\"]/span[@class=\"endorse-item-name-text\"]", true);
                                hscList.Add(hscName + "(" + hscCount + ");");
                            }
                            hsc = CommonUtil.StringArrayToString(hscList.ToArray(), "");
                        }

                        //个人信息
                        HtmlNodeCollection grxxNodes = documentNode.SelectNodes("//li[@class=\"personal-info\"]/table[@id=\"personal-info-view\"]/tbody/tr");
                        if (grxxNodes != null)
                        {
                            List<String> grxxList = new List<string>();
                            foreach (HtmlNode grxxNode in grxxNodes)
                            {
                                string grxxName = HtmlDocumentHelper.TryGetNodeInnerText(grxxNode, "./th", true);
                                string grxxValue = HtmlDocumentHelper.TryGetNodeInnerText(grxxNode, "./td", true);
                                grxxList.Add(grxxName + ":" + grxxValue + "; ");
                            }
                            grxx = CommonUtil.StringArrayToString(grxxList.ToArray(), "");
                        }

                        //联系详情
                        lxxq = HtmlDocumentHelper.JointNodeInnerText(documentNode, "//li[@id=\"contact-comments\"]/div[@id=\"contact-comments-view\"]/p", "; ", true, true, null, null);

                        //参与组织
                        HtmlNodeCollection cyzzNodes = documentNode.SelectNodes("//div[@id=\"background-organizations\"]/div/div");
                        if (cyzzNodes != null)
                        {
                            List<String> cyzzList = new List<string>();
                            foreach (HtmlNode cyzzNode in cyzzNodes)
                            {
                                string cyzzName = HtmlDocumentHelper.TryGetNodeInnerText(cyzzNode, "./hgroup/h4", true);
                                string cyzzTime = HtmlDocumentHelper.JointNodeInnerText(cyzzNode, "./span[@class=\"organizations-date\"]/*", "", true, true, "", " ");
                                cyzzList.Add(cyzzName + ", " + cyzzTime + ";");
                            }
                            cyzz = CommonUtil.StringArrayToString(cyzzList.ToArray(), "\r\n");
                        }

                        //所做项目
                        HtmlNodeCollection szxmNodes = documentNode.SelectNodes("//div[@id=\"background-projects\"]/div/div");
                        if (szxmNodes != null)
                        {
                            List<String> szxmList = new List<string>();
                            foreach (HtmlNode szxmNode in szxmNodes)
                            {
                                string szxmName = HtmlDocumentHelper.TryGetNodeInnerText(szxmNode, "./hgroup/h4", true);
                                string szxmTime = HtmlDocumentHelper.JointNodeInnerText(szxmNode, "./span[@class=\"projects-date\"]/*", "", true, true, "", " ");
                                szxmList.Add(szxmName + ", " + szxmTime + ";");
                            }
                            szxm = CommonUtil.StringArrayToString(szxmList.ToArray(), "\r\n");
                        }

                        //推荐信
                        //直接全部文本摘录下来吗?????????????????????????这里只是爬取了推荐信的个数
                        tjx = HtmlDocumentHelper.JointNodeInnerText(documentNode, "//div[@id=\"endorsements\"]/ul[@class=\"endorsements-nav\"]/li/a", "; ", true, true, null, null);

                        //联系人, 这里只是爬取了几个共同联系人
                        lxr = HtmlDocumentHelper.TryGetNodeInnerText(documentNode, "//div[@id=\"connections\"]/div[@id=\"connections-view\"]/ul[@class=\"connections-nav\"]/li[@class=\"nav-shared\"]", true, true, null, null);

                        //关注内容
                        HtmlNodeCollection gznrNodes = documentNode.SelectNodes("//div[@id=\"following-container\"]/div[@class=\"profile-following\"]/div/div");
                        if (gznrNodes != null)
                        {
                            List<String> gznrList = new List<string>();
                            foreach (HtmlNode gznrNode in gznrNodes)
                            {
                                string gznrCategory = HtmlDocumentHelper.TryGetNodeInnerText(gznrNode, "./h3", true);
                                HtmlNodeCollection fNodes = gznrNode.SelectNodes("./ul/li");
                                List<String> fList = new List<string>();
                                if (fNodes != null)
                                {
                                    foreach (HtmlNode fNode in fNodes)
                                    {
                                        string fStr = HtmlDocumentHelper.JointNodeInnerText(fNode, "./p", " ", true, true, null, ", ");
                                        if (fStr != null && fStr.Length > 0)
                                        {
                                            fStr = fStr.Substring(0, fStr.Length - 1);
                                            fList.Add(fStr);
                                        }
                                    }
                                    gznrList.Add(gznrCategory + ": " + CommonUtil.StringArrayToString(fList.ToArray(), "; "));
                                }
                            }
                            gznr = CommonUtil.StringArrayToString(gznrList.ToArray(), "\r\n");
                        }
                        #endregion
                    }

                    Dictionary<string, string> f2vs = new Dictionary<string, string>();
                    f2vs.Add("姓名", xm);
                    f2vs.Add("英文名", englishName);
                    f2vs.Add("目前就职", mqjz);
                    f2vs.Add("职位", zw);
                    f2vs.Add("所在地", szd);
                    f2vs.Add("曾经就职", cjjz);
                    f2vs.Add("教育经历", jyjl);
                    f2vs.Add("电话", dh);
                    f2vs.Add("邮箱", yx);
                    f2vs.Add("微信", wx);
                    f2vs.Add("微信二维码", wxewm);
                    f2vs.Add("网站", wz);
                    f2vs.Add("工作经历", gzjl);
                    f2vs.Add("资格认证", zgrz);
                    f2vs.Add("支持的组织机构", zcdzzjg);
                    f2vs.Add("出版作品", cbzp);
                    f2vs.Add("语言能力", yynl);
                    f2vs.Add("技能", jn);
                    f2vs.Add("还擅长", hsc);
                    f2vs.Add("个人信息", grxx);
                    f2vs.Add("联系详情", lxxq);
                    f2vs.Add("参与组织", cyzz);
                    f2vs.Add("所做项目", szxm);
                    f2vs.Add("推荐信", tjx);
                    f2vs.Add("联系人", lxr);
                    f2vs.Add("关注内容", gznr);
                    f2vs.Add("站内网址", znwz);
                    return f2vs;
                }
                else
                {
                    return null;
                }
            }
            else
            {
                return null;
            }
        }
        #endregion
    }
}
