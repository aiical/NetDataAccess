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
using NetDataAccess.AppAccessBase;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Android;
using System.Collections.ObjectModel;
using System.Xml;

namespace NetDataAccess.Extended.Linkedin.KeyWordsSearchAtLinkedinApp
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

        private string _LogoutPageUrl = null;
        private string LogoutPageUrl
        {
            get
            {
                return _LogoutPageUrl;
            }
            set
            {
                _LogoutPageUrl = value;
            }
        }

        private string _LogoutSucceedCheckUrl = null;
        private string LogoutSucceedCheckUrl
        {
            get
            {
                return _LogoutSucceedCheckUrl;
            }
            set
            {
                _LogoutSucceedCheckUrl = value;
            }
        }

        public override bool BeforeAllGrab()
        {
            string[] ps = this.Parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            this.LinkedinLoginPageUrl = ps[0];
            this.LinkedinLoginSucceedCheckUrl = ps[1];
            this.LogoutPageUrl = ps[2];
            this.LogoutSucceedCheckUrl = ps[3];
            return true;
        }

        /// <summary>
        /// 初始化与手机App的连接
        /// </summary>
        private AndroidAppAccess InitAppAccess()
        {
            AndroidAppAccess appAccess = new AndroidAppAccess();
            Dictionary<string, string> initParams = new Dictionary<string, string>();
            initParams.Add("deviceName", "HUAWEI G700-U00");
            initParams.Add("platformVersion", "4.2");
            initParams.Add("appPackage", "com.linkedin.android");
            initParams.Add("appActivity", "com.linkedin.android.authenticator.LaunchActivity"); ;
            initParams.Add("url", "http://127.0.0.1:4723/wd/hub");
            appAccess.InitDriver(initParams);

            //留出五秒钟的实际，等待手机被唤醒，及app被启动
            Thread.Sleep(5000);

            return appAccess;
        }

        private void CloseAppAccess(AndroidAppAccess appAccess)
        {
            if (appAccess != null)
            {
                appAccess.Close();
            }
        }

        /// <summary>
        /// 判断是否需要登出
        /// </summary>
        /// <param name="appAccess"></param>
        private bool CheckNeedLogout(AndroidAppAccess appAccess)
        {
            AndroidElement navElement = this.GetNavElement(appAccess, false);
            //判断存在“立即加入”和“登录”按钮
            bool hasLogout = appAccess.CheckCurrentPageContainText(new string[] { "立即加入", "登录" }, true);
            if (navElement == null && !hasLogout)
            {
                //可能是出现了广告窗口，那么back关闭广告
                appAccess.ClickBackButton();
                Thread.Sleep(5000);
                //判断存在“立即加入”和“登录”按钮
                return !appAccess.CheckCurrentPageContainText(new string[] { "立即加入", "登录" }, true);
            }
            else
            {
                return !hasLogout;
            }
        }

        private AndroidElement GetNavElement(AndroidAppAccess appAccess, bool errorNone)
        {
            ReadOnlyCollection< AndroidElement> navElements = appAccess.GetElementsByXPath("//android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.FrameLayout/android.support.v4.widget.DrawerLayout/android.view.View/android.widget.LinearLayout/android.widget.LinearLayout", errorNone);
            if (navElements == null)
            {
                return null;
            }
            else
            {
                return navElements[navElements.Count - 1];
            }
        }

        /// <summary>
        /// 执行登出
        /// </summary>
        /// <param name="appAccess"></param>
        private void DoLogout(AndroidAppAccess appAccess)
        {
            AndroidElement navElement = GetNavElement(appAccess, true);
            AndroidElement myElement = appAccess.GetElementByClassNameAndText("android.widget.TextView", "我", true, true, true);
            myElement.Click();

            //等待“我”的页面展示出来
            AndroidElement checkUserElment = appAccess.GetElementByClassNameAndText("android.widget.TextView", new string[] { "谁看过我" }, true, false);
            if (checkUserElment != null)
            {
                appAccess.Swipe(new Point(200, 800), new Point(200, 100), 1000);
                AndroidElement personalSettingElement = appAccess.GetElementByClassNameAndText("android.widget.TextView", new string[] { "个人设置" }, true, true);
                personalSettingElement.Click();
            }
            else
            {
                try
                {
                    AndroidElement toolElement = appAccess.GetElementByXPath("//android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.FrameLayout/android.support.v4.widget.DrawerLayout/android.view.View/android.widget.LinearLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.ImageView", true);
                    toolElement.Click();
                }
                catch (Exception ex)
                {
                    throw new Exception("无法打开个人设置页面", ex);
                }
            }

            appAccess.GetElementByClassNameAndText("android.widget.TextView", new string[] { "邮箱地址" }, true, true);
            appAccess.Swipe(new Point(200, 800), new Point(200, 100), 1000);
            AndroidElement logoutElement = appAccess.GetElementByClassNameAndText("android.widget.TextView", new string[] { "退出帐号" }, true, true);
            logoutElement.Click();

            appAccess.CheckCurrentPageContainText(new string[] { "立即加入", "登录" }, true, true);
        }

        /// <summary>
        /// 执行登录
        /// </summary>
        /// <param name="appAccess"></param>
        private void DoLogin(AndroidAppAccess appAccess, string loginName, string loginPassword)
        {
            AndroidElement toLoginButton = appAccess.GetElementByClassNameAndText("android.widget.Button", new string[] { "登录" }, true, true);
            toLoginButton.Click();
            appAccess.GetElementByClassNameAndText("android.widget.Button", new string[] { "忘记密码？" }, true, true);

            AndroidElement loginNameLayoutElement = appAccess.GetElementByClassNameAndText("android.widget.LinearLayout", new string[] { "邮箱或电话" }, true, true);
            AndroidElement loginNameElement = appAccess.GetElementByClassNameAndIndex(loginNameLayoutElement, "android.widget.EditText", 0, true);
            loginNameElement.Click();
            loginNameElement.Clear();
            loginNameElement.SendKeys(loginName);

            AndroidElement loginPasswordLayoutElement = appAccess.GetElementByClassNameAndText("android.widget.LinearLayout", new string[] { "密码" }, true, true);
            AndroidElement loginPasswordElement = appAccess.GetElementByClassNameAndIndex(loginPasswordLayoutElement, "android.widget.EditText", 0, true);
            loginPasswordElement.Click();
            loginPasswordElement.Clear();
            loginPasswordElement.SendKeys(loginPassword);

            AndroidElement loginButton = appAccess.GetElementByClassNameAndText("android.widget.Button", new string[] { "登录" }, true, true);
            loginButton.Click();

            //判断是否登录成功
            this.GetNavElement(appAccess, true);
        }

        /// <summary>
        /// 使用某关键词查询
        /// </summary>
        /// <param name="appAccess"></param>
        private List<Dictionary<string, string>> SearchPersonByKeyWord(AndroidAppAccess appAccess, string keyWords)
        {
            //存储查询到的人员类别
            List<Dictionary<string, string>> personList = new List<Dictionary<string, string>>();

            AndroidElement searchElement = appAccess.GetElementByClassNameAndText("android.widget.TextView", new string[] { "搜索用户、职位及其他内容" }, true, true);
            searchElement.Click();
            Thread.Sleep(5000);
            searchElement.SendKeys(keyWords);
            Thread.Sleep(5000);
            appAccess.SendKeyEvent(66);
            Thread.Sleep(5000);

            Dictionary<string, string> personSimpleInfos = new Dictionary<string, string>();
            AndroidElement personListElement = null;
            try
            {
                personListElement = this.GetNextPersonElement(appAccess, personSimpleInfos, 0);
            }
            catch (Exception ex)
            {
                throw new Exception("在app中根据关键字查询，获取查询结果列表中第一位会员信息出错", ex);
            }

            Dictionary<string, string> personUrlDic = new Dictionary<string, string>();
            while (personListElement != null)
            {
                //点击列表项，跳转到个人页面，记录个人信息
                string personName = this.GetNameByListPersonElement(appAccess, personListElement);
                try
                {
                    string personUrl = this.GetPersonLinkedinUrlFromPersonPage(appAccess, personListElement, keyWords);

                    if (personUrl == null)
                    {
                        this.RunPage.InvokeAppendLogText("无法从app的个人页面里找到linkedin个人网址", LogLevelType.System, true);
                    }
                    else
                    {
                        if (!personUrlDic.ContainsKey(personUrl))
                        {
                            personUrlDic.Add(personUrl, null);

                            Dictionary<string, string> personInfo = new Dictionary<string, string>();
                            personInfo.Add("personName", personName);
                            personInfo.Add("personUrl", personUrl);
                            personList.Add(personInfo);
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception("使用APP抓取的个人信息出错, personName = " + personName, ex);
                }

                try
                {
                    personListElement = this.GetNextPersonElement(appAccess, personSimpleInfos, 0);
                }
                catch (Exception ex)
                {
                    throw new Exception("获取下一位会员信息出错, personName = " + personName, ex);
                }
            }
            return personList;
        }

        private string GetPersonLinkedinUrlFromPersonPage(AndroidAppAccess appAccess, AndroidElement personListElement,string keyWords)
        {
            string url = null;
            string elementText = this.GetListPersonElementText(appAccess, personListElement);

            string personUrlName = keyWords + "_" + elementText;
            string personUrlFilePath = this.RunPage.GetFilePath(personUrlName, this.RunPage.GetDetailSourceFileDir());
            if (File.Exists(personUrlFilePath))
            {
                url = FileHelper.GetTextFromFile(personUrlFilePath);
                return url == null || url.Length == 0 ? null : url;
            }
            else
            {
                personListElement.Click();
                try
                {
                    url = this.GetPersonLinkedinUrlFromElement(appAccess, 0);
                }
                catch (Exception ex)
                {
                    this.RunPage.InvokeAppendLogText("从APP个人页面里获取个人页面Url失败. " + ex.Message, LogLevelType.Error, true);
                }
                appAccess.ClickBackButton();
                FileHelper.SaveTextToFile(url, personUrlFilePath);
                return url;
            }
        }

        private bool CheckIsAnonymousUserPage(AndroidAppAccess appAccess)
        { 
            int checkCount = 0;
            while (checkCount < 5)
            {
                try
                {
                    return appAccess.CheckContainTextByXPath("//android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.FrameLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.TextView", new string[] { "该会员在您的人脉圈外" }, true);
                }
                catch (Exception ex)
                {
                    checkCount++;
                    Thread.Sleep(2000);
                }
            } 
                         
            this.RunPage.InvokeAppendLogText("无法判断是否是匿名用户，即检测'该会员在您的人脉圈外'是否存在不成功", LogLevelType.Error, true);
            return true;
        }

        private string GetPersonLinkedinUrlFromElement(AndroidAppAccess appAccess, int swipeCount)
        {
            Thread.Sleep(2000); 
            bool canRead = !this.CheckIsAnonymousUserPage(appAccess); 

            if (canRead)
            {
                XmlElement rootXmlElement = appAccess.GetXmlRootElement();
                XmlNode containerXmlElement = rootXmlElement.SelectSingleNode("//android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.FrameLayout/android.view.View/android.view.View/android.support.v7.widget.RecyclerView");

                XmlNodeList infoSectionXmlElements = containerXmlElement.SelectNodes("./android.widget.LinearLayout");
                foreach (XmlNode infoSectionXmlElement in infoSectionXmlElements)
                {
                    XmlNode sectionNameXmlElement = infoSectionXmlElement.SelectSingleNode("./android.widget.FrameLayout/android.widget.LinearLayout/android.widget.LinearLayout[1]/android.widget.TextView");
                    if (sectionNameXmlElement != null && sectionNameXmlElement.Attributes["text"].Value == "联系信息")
                    {
                        XmlNode linkElement = infoSectionXmlElement.SelectSingleNode("./android.widget.FrameLayout/android.widget.LinearLayout/android.widget.LinearLayout[2]/android.widget.RelativeLayout/android.widget.LinearLayout/android.widget.LinearLayout/android.widget.TextView[2]");
                        if (linkElement != null && linkElement.Attributes["text"].Value.Contains(".linkedin.com/"))
                        {
                            return linkElement.Attributes["text"].Value.Trim();
                        }
                    }
                }

                if (swipeCount < 10)
                {
                    swipeCount++;
                    appAccess.Swipe(new Point(200, 900), new Point(200, 200), 1000);
                    return this.GetPersonLinkedinUrlFromElement(appAccess, swipeCount);
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

        private string GetListPersonElementText(AndroidAppAccess appAccess, AndroidElement listLElement)
        {
            try
            {
                StringBuilder textBuilder = new StringBuilder();
                AppiumWebElement nameLElement = listLElement.FindElementByClassName("android.widget.LinearLayout");
                if (nameLElement != null)
                {
                    AppiumWebElement nameElement = nameLElement.FindElementByClassName("android.widget.TextView");
                    if (nameElement != null)
                    {
                        ReadOnlyCollection<AppiumWebElement> propertyElements = listLElement.FindElementsByClassName("android.widget.TextView");
                        for (int i = 0; i < propertyElements.Count; i++)
                        {
                            textBuilder.AppendLine();
                            textBuilder.Append(propertyElements[i].Text.Trim());
                        }
                        return textBuilder.ToString();
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                //总是莫名其妙报错
                return null;
            }
        }

        private string GetNameByListPersonElement(AndroidAppAccess appAccess, AndroidElement listLElement)
        {
            AppiumWebElement nameElement = listLElement.FindElementByClassName("android.widget.LinearLayout").FindElementByClassName("android.widget.TextView");
            return nameElement == null ? null : nameElement.Text;
        }

        private AndroidElement GetNextPersonElement(AndroidAppAccess appAccess, Dictionary<string, string> personSimpleInfos, int noneNewCount)
        {
            AndroidElement listContainerElement = appAccess.GetElementByClassNameAndIndex("android.support.v7.widget.RecyclerView", 0, true);
            if (!appAccess.CheckElementContainText(listContainerElement, new string[] { "没有找到结果" }, true))
            {
                ReadOnlyCollection<AndroidElement> listElements = appAccess.GetElementsByClassName(listContainerElement, "android.widget.FrameLayout", false);
                if (listElements != null)
                {
                    for (int i = 0; i < listElements.Count; i++)
                    {
                        AppiumWebElement listRElement = listElements[i].FindElementByClassName("android.widget.RelativeLayout");
                        if (listRElement != null)
                        {
                            ReadOnlyCollection<AppiumWebElement> listLElements = listRElement.FindElementsByClassName("android.widget.LinearLayout");
                            if (listLElements != null && listLElements.Count > 0)
                            {
                                string elementText = this.GetListPersonElementText(appAccess, (AndroidElement)listLElements[0]);
                                if (elementText == null)
                                {
                                    //获取个人信息出错，可能是个人信息展示不完全造成的，那么翻页吧
                                    break;
                                }
                                else if (elementText.Contains("领英会员"))
                                {
                                    continue;
                                }
                                else if (!personSimpleInfos.ContainsKey(elementText))
                                {
                                    personSimpleInfos.Add(elementText, null);
                                    return (AndroidElement)listLElements[0];
                                }
                            }
                        }
                    }

                    if (noneNewCount < 3)
                    {
                        noneNewCount++;
                        //滑动翻页 
                        appAccess.Swipe(new Point(100, 700), new Point(100, 200), 500);
                        Thread.Sleep(500);
                        return this.GetNextPersonElement(appAccess, personSimpleInfos, noneNewCount);
                    }
                    else
                    {
                        //滑动翻页了3次，仍然没有新的，那么就认为真的没有新的了
                        return null;
                    }
                }
                else
                {
                    //没有匹配的结果
                    return null;
                }
            }
            else
            {
                //没有匹配的结果
                return null;
            }
        }

        public override void GetDataByOtherAcessType(Dictionary<string, string> listRow)
        {

            string keyWords = listRow["keyWords"];
            string loginName = listRow["loginName"];
            string loginPassword = listRow["loginPassword"];

            string fileName = "_" + keyWords + "_" + loginName;
            List<Dictionary<string, string>> personInfos = this.RunPage.TryGetInfoFromMiddleFile(fileName, new string[] { "personName", "personUrl" });
            if (personInfos == null)
            {
                this.RunPage.InvokeAppendLogText("开始使用APP搜索相关人员， 关键词为'" + keyWords + "'", LogLevelType.System, true);

                AndroidAppAccess appAccess = null;

                try
                {
                    this.RunPage.InvokeAppendLogText("开始连接手机APP", LogLevelType.System, true);
                    appAccess = this.InitAppAccess();
                    this.RunPage.InvokeAppendLogText("连接手机APP成功", LogLevelType.System, true);

                    bool needLogout = false;

                    try
                    {
                        this.RunPage.InvokeAppendLogText("检测是否需要登出", LogLevelType.System, true);
                        needLogout = this.CheckNeedLogout(appAccess);
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("检查是否需要登出出错.", ex);
                    }
                    if (needLogout)
                    {
                        try
                        {
                            this.RunPage.InvokeAppendLogText("开始登出", LogLevelType.System, true);
                            this.DoLogout(appAccess);
                            this.RunPage.InvokeAppendLogText("已经登出", LogLevelType.System, true);
                        }
                        catch (Exception ex)
                        {
                            throw new Exception("登出失败.", ex);
                        }
                    }

                    try
                    {
                        this.RunPage.InvokeAppendLogText("开始登录", LogLevelType.System, true);
                        this.DoLogin(appAccess, loginName, loginPassword);
                        this.RunPage.InvokeAppendLogText("已经登录", LogLevelType.System, true);
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("登录失败.", ex);
                    }

                    try
                    {
                        personInfos = this.SearchPersonByKeyWord(appAccess, keyWords);
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("在app中根据关键字抓取出错, keyWords = " + keyWords);
                    }

                    this.RunPage.SaveInfoToMiddleFile(fileName, new string[] { "personName", "personUrl" }, personInfos);
                    this.RunPage.InvokeAppendLogText("完成使用手机APP搜索关键词'"+keyWords+"", LogLevelType.System, true);
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                finally
                {
                    if (appAccess != null)
                    {
                        this.CloseAppAccess(appAccess);
                    }
                }
            }

            try
            {
                string checkedRelatedPersonInfosResultFilePath = this.RunPage.GetFilePath("_SearchResult_LinkedinApp_" + loginName + "_" + keyWords + ".xlsx", this.RunPage.GetExportDir());
                if (!File.Exists(checkedRelatedPersonInfosResultFilePath))
                {
                    this.RunPage.InvokeAppendLogText("登录Linkedin系统", LogLevelType.System, true);
                    LoginLinkedin.LoginByRandomUser(this.RunPage, this.LinkedinLoginPageUrl, this.LinkedinLoginSucceedCheckUrl);
                    this.RunPage.InvokeAppendLogText("已登录Linkedin系统", LogLevelType.System, true);

                    this.RunPage.InvokeAppendLogText("根据个人页面地址（人员列表是从手机APP搜索到的，关键词为'" + keyWords + "'），从Linkedin网页版获取个人信息", LogLevelType.System, true);
                    string personPageInfosWithJustDownloadMarkFileName = "_" + loginName + "_" + keyWords + "_personPageInfosWithJustDownloadMark";
                    //读取历史生成的个人网页网址
                    List<Dictionary<string, string>> allPersonPageInfosWithJustDownloadMark = this.RunPage.TryGetInfoFromMiddleFile(personPageInfosWithJustDownloadMarkFileName, new string[] { "personUrl", "personName", "isJustDownload" });
                    if (allPersonPageInfosWithJustDownloadMark == null)
                    {
                        allPersonPageInfosWithJustDownloadMark = ProcessPersonPage.GetAllPersonPages(this.RunPage, personInfos, loginName, loginPassword);
                        this.RunPage.SaveInfoToMiddleFile(personPageInfosWithJustDownloadMarkFileName, new string[] { "personUrl", "personName", "isJustDownload" }, allPersonPageInfosWithJustDownloadMark);
                    }

                    string personInfosFilePath = this.RunPage.GetFilePath("_SearchResult_LinkedinApp_" + loginName + "_" + keyWords + ".xlsx", this.RunPage.GetExportDir());
                    if (!File.Exists(personInfosFilePath))
                    {
                        List<Dictionary<string, string>> personInfoList = ProcessPersonPage.GetPersonInfoFromLocalPages(this.RunPage, allPersonPageInfosWithJustDownloadMark, false, null);
                        ProcessPersonPage.SavePersonInfoToFile(this.RunPage, personInfoList, personInfosFilePath);
                    }
                    this.RunPage.InvokeAppendLogText("完成获取并处理个人页面（人员列表是从手机APP搜索到的，关键词为'" + keyWords + "'）", LogLevelType.System, true);
                     
                    /*
                    this.RunPage.InvokeAppendLogText("从已爬取到的页面中，找到'看过本页的会员还看了'栏目里的个人信息，递归获取", LogLevelType.System, true);
                    string checkedRelatedPersonIdsFileName = "_" + keyWords + "_" + loginName + "_CheckedRelated";
                    List<Dictionary<string, string>> allCheckedRelatedIds = this.RunPage.TryGetInfoFromMiddleFile(checkedRelatedPersonIdsFileName, new string[] { "checkedRelatedPersonInfoId", "levelCount" });
                    if (allCheckedRelatedIds == null)
                    {
                        allCheckedRelatedIds = new List<Dictionary<string, string>>();
                        Dictionary<string, string> allRelatedPersonUrlInfos = new Dictionary<string, string>();
                        ProcessPersonPage.GetRelatedPersonInfos(this.RunPage, allPersonPageInfosWithJustDownloadMark, allRelatedPersonUrlInfos, allCheckedRelatedIds, keyWords);
                        this.RunPage.SaveInfoToMiddleFile(checkedRelatedPersonIdsFileName, new string[] { "checkedRelatedPersonInfoId", "levelCount" }, allCheckedRelatedIds);
                    }

                    this.RunPage.InvokeAppendLogText("从已爬取到的页面中，找到'看过本页的会员还看了'栏目里的个人信息", LogLevelType.System, true);
                    string checkedRelatedPersonUrlsFileName = "_" + loginName + "_" + keyWords + "_personUrl_CheckedRelated";
                    List<Dictionary<string, string>> allCheckAllCheckedRelatedUrls = this.RunPage.TryGetInfoFromMiddleFile(checkedRelatedPersonUrlsFileName, new string[] { "personUrl", "personName" });
                    if (allCheckAllCheckedRelatedUrls == null)
                    {
                        allCheckAllCheckedRelatedUrls = ProcessPersonPage.GetAllPersonPages(this.RunPage, allCheckedRelatedIds);
                    }

                    List<string> checkRelatedPersonPageIds = new List<string>();
                    foreach (Dictionary<string, string> checkedRelatedId in allCheckedRelatedIds)
                    {
                        string checkedRelatedPersonInfoId = checkedRelatedId["checkedRelatedPersonInfoId"];
                        checkRelatedPersonPageIds.Add(checkedRelatedPersonInfoId);
                    }
                    List<Dictionary<string, string>> relatedPersonInfoList = ProcessPersonPage.GetPersonInfoFromLocalPages(this.RunPage, allCheckAllCheckedRelatedUrls, false, null);
                    ProcessPersonPage.SavePersonInfoToFile(this.RunPage, relatedPersonInfoList, checkedRelatedPersonInfosResultFilePath);
                    */

                    this.RunPage.InvokeAppendLogText("完成递归爬取到关键词'" + keyWords + "'相关的所有的'看过本页的会员还看了'", LogLevelType.System, true);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("通过网页获取/处理个人信息出错", ex);
            }
        }
    }
}