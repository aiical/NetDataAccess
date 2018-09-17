using NetDataAccess.Base.Browser;
using NetDataAccess.Base.Common;
using NetDataAccess.Base.Config;
using NetDataAccess.Base.DB;
using NetDataAccess.Base.EnumTypes;
using NetDataAccess.Base.UI;
using NetDataAccess.Base.Web;
using NetDataAccess.Base.Writer;
using NetDataAccess.Extended.Linkedin.Common;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace NetDataAccess.Extended.Linkedin.KeyWordsSearchAtLinkedin
{
    [System.Runtime.InteropServices.ComVisibleAttribute(true)]
    public class GetKeywordsLinkedinDetailPage : ExternalRunWebPage
    {
        public override void WebBrowserHtml_AfterDoNavigate(string pageUrl, Dictionary<string, string> listRow, string tabName)
        {
            this.RunPage.ShowTabPage(tabName);
        }

        public override void WebBrowserHtml_AfterPageLoaded(string pageUrl, Dictionary<string, string> listRow, IWebBrowser webBrowser)
        {
            Thread.Sleep(2000);
            string checkValue = "pv-top-card-v2-section__info mr5";
            string webBrowserUrl = this.RunPage.InvokeGetWebBrowserPageUrl(webBrowser);

            string webText = this.RunPage.InvokeGetPageHtml(webBrowser).ToLower();
            if (!webText.Contains(checkValue))
            {
                throw new Exception("页面加载地址错误, webBrowserUrl=" + webBrowserUrl + ", pageUrl=" + pageUrl);
            }
            ProcessWebBrowser.AutoScroll(this.RunPage, webBrowser, 3000, 500, 1000, 2000);

            this.ClickAllMoreLinks(webBrowser);

            Random r = new Random(DateTime.Now.Millisecond);
            Thread.Sleep(r.Next(10) * 1000);
        }

        private void ClickAllMoreLinks(IWebBrowser webBrowser)
        {
            AddClickMoreMethod(webBrowser);
        }
         
        public override void BeforeGrabOne(string pageUrl, Dictionary<string, string> listRow, bool existLocalFile)
        {
            base.BeforeGrabOne(pageUrl, listRow, existLocalFile);
            if (this.CheckNeedLogin())
            {
                if (this.CheckCanLogin())
                {
                    this.Login();
                }
                else
                {
                    Thread.Sleep(1000 * 60 * 10);
                    throw new Exception("没有可用的账号");
                }
            }
        }

        private int _CurrentUserIndex = -1;
        private int _CurrentUserRequestCount = 0;

        private bool CheckNeedLogin()
        {
            if (_CurrentUserIndex < 0)
            {
                this.Logout();
                _CurrentUserIndex++;
                _CurrentUserRequestCount = 0;
                return true;
            }
            else if (_CurrentUserIndex >= _UserInfoList.Count)
            {
                return true;
            }
            else
            {
                string[] userInfo = _UserInfoList[_CurrentUserIndex];
                int requestCountLimit = int.Parse(userInfo[2]);
                if (_CurrentUserRequestCount > requestCountLimit)
                {
                    this.Logout();
                    _CurrentUserIndex++;
                    _CurrentUserRequestCount = 0;
                    return true;
                }
                else
                {
                    _CurrentUserRequestCount++;
                    return false;
                }
            }
        }

        private bool CheckCanLogin()
        {
            if (_CurrentUserIndex < _UserInfoList.Count)
            {
                return true;
            }
            else if(DateTime.Now.Hour == 13)
            {
                //现在是下午1点半到2点之间，那么重新启动爬取
                _CurrentUserIndex = 0;
                return true;
            }
            else
            {
                return false;
            }
        }

        private void Login()
        {
            string pageUrl = "https://www.linkedin.com";
            string tabName = "login";
            IWebBrowser webBrowser = this.RunPage.InvokeShowWebPage(pageUrl, tabName, WebBrowserType.Chromium, false);
            this.RunPage.ShowTabPage(tabName);
            string htmlContent = null;
            int waitCount = 0;
            int timeout = 30000;
            while (htmlContent == null)
            {
                if (SysConfig.WebPageRequestInterval * waitCount > timeout)
                {
                    //超时
                    throw new GrabRequestException("请求Logout页超时. PageUrl = " + pageUrl);
                }
                //等待
                waitCount++;
                Thread.Sleep(SysConfig.WebPageRequestInterval);
                htmlContent = webBrowser.GetDocumentText();
            } 

            string[] userInfo = _UserInfoList[_CurrentUserIndex];

            string userName = userInfo[0].Trim();
            string password = userInfo[1].Trim();

            string inputUserInfoCode = "document.getElementById('login-email').click();document.getElementById('login-email').value = '" + userName + "';document.getElementById('login-password').click();document.getElementById('login-password').value = '" + password + "';document.getElementById('login-submit').disabled=false;";
            webBrowser.AddScriptMethod(inputUserInfoCode);
            Thread.Sleep(3000);

            string submitUserInfoCode = "document.getElementById('login-submit').click();";
            webBrowser.AddScriptMethod(submitUserInfoCode);
            Thread.Sleep(3000);
        }

        private void Logout()
        {
            string pageUrl = "http://www.linkedin.com/m/logout";
            string tabName = "logout";
            IWebBrowser webBrowser = this.RunPage.InvokeShowWebPage(pageUrl, tabName, WebBrowserType.Chromium, false);
            this.RunPage.ShowTabPage(tabName);
            string htmlContent = null;
            int waitCount = 0;
            int timeout = 30000;
            while (htmlContent == null)
            {
                if (SysConfig.WebPageRequestInterval * waitCount > timeout)
                {
                    //超时
                    throw new GrabRequestException("请求Logout页超时. PageUrl = " + pageUrl);
                }
                //等待
                waitCount++;
                Thread.Sleep(SysConfig.WebPageRequestInterval);
                htmlContent = webBrowser.GetDocumentText();
            }
        }

        private List<string[]> _UserInfoList = null;

        public override bool BeforeAllGrab()
        {
            _UserInfoList = new List<string[]>();
            string[] userInfoStrs = this.Parameters.Split(new string[] { ";;;" }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < userInfoStrs.Length; i++)
            {
                string userInfoStr = userInfoStrs[i];

                string[] userInfo = userInfoStr.Split(new string[] { "|||" }, StringSplitOptions.RemoveEmptyEntries);
                _UserInfoList.Add(userInfo);
            }
            return true;
        }



        private void AddClickMoreMethod(IWebBrowser webBrowser)
        {
            string scriptMethodCodeA = "$('button[class=\"pv-profile-section__card-action-bar pv-skills-section__additional-skills artdeco-container-card-action-bar\"]').click();" ;
            string scriptMethodCodeB = "$('button[class=\"pv-profile-section__see-more-inline pv-profile-section__text-truncate-toggle link\"]').click();";

            this.RunPage.InvokeAddScriptMethod(webBrowser, scriptMethodCodeA);
            Thread.Sleep(3000);
            this.RunPage.InvokeAddScriptMethod(webBrowser, scriptMethodCodeB); 
            Thread.Sleep(3000);
        }


        public override bool AfterAllGrab(IListSheet listSheet)
        {
            //this.GetUserInfoInPages(listSheet);
            return true;
        }

        private void GetUserInfoInPages(IListSheet listSheet)
        {
            int rowCount = listSheet.GetListDBRowCount();
            for (int i = 0; i < rowCount; i++)
            {

            }
        }

        private ExcelWriter GetExcelBaseWriter()
        {
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                "url",
                "id",
                "名字", 
                "目前工作", 
                "地区", 
                "keywords"});
            string resultFilePath = Path.Combine(exportDir, "Linkedin个人基本信息.xlsx");

            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic);
            return resultEW;
        }

        private ExcelWriter GetExcelExpWriter()
        {
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                "id",
                "名字", 
                "岗位", 
                "公司/组织", 
                "任职时间", 
                "任职时长", 
                "所在地区", 
                "描述"});
            string resultFilePath = Path.Combine(exportDir, "Linkedin个人工作经验.xlsx");

            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic);
            return resultEW;
        }

        private ExcelWriter GetExcelEduWriter()
        {
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                "id",
                "名字", 
                "学校",  
                "学习时间",   
                "描述"});
            string resultFilePath = Path.Combine(exportDir, "Linkedin个人教育经历.xlsx");

            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic);
            return resultEW;
        }

        private ExcelWriter GetExcelSkillWriter()
        {
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                "id",
                "名字", 
                "技能",  
                "学习时间",   
                "描述"});
            string resultFilePath = Path.Combine(exportDir, "Linkedin个人技能.xlsx");

            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic);
            return resultEW;
        }

    }
}
