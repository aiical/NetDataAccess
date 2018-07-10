using NetDataAccess.Base.Config;
using NetDataAccess.Base.Reader;
using NetDataAccess.Base.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NetDataAccess.Extended.Linkedin.Common
{
    public class LoginLinkedin
    {
        /// <summary>
        /// 从输入的任务文件里，获取用户登录信息
        /// </summary>
        /// <param name="runPage"></param>
        /// <param name="loginPageUrl"></param>
        /// <param name="loginSucceedCheckUrl"></param>
        /// <returns></returns>
        public static bool Login(IRunWebPage runPage, string loginPageUrl, string loginSucceedCheckUrl)
        {
            Dictionary<string, string> seedRow = GetRowInExcelList(runPage);
            if (!seedRow.ContainsKey("loginName") || !seedRow.ContainsKey("loginPassword"))
            {
                throw new Exception("输入的任务文件里没有指定Linkedin的用户loginName和密码loginPassword信息");
            }
            string loginName = seedRow["loginName"];

            string loginPassword = seedRow["loginPassword"];

            return Login(runPage, loginName, loginPassword, loginPageUrl, loginSucceedCheckUrl);
        }

        /// <summary>
        /// 随机从公共账号池里获取用户登录信息
        /// </summary>
        /// <param name="runPage"></param>
        /// <param name="loginPageUrl"></param>
        /// <param name="loginSucceedCheckUrl"></param>
        /// <returns></returns>
        public static bool LoginByRandomUser(IRunWebPage runPage, string loginPageUrl, string loginSucceedCheckUrl)
        {
            Dictionary<string, string> userInfo = ProcessLinkedinUser.GetUserLoginInfo();

            if (!userInfo.ContainsKey("loginName") || !userInfo.ContainsKey("loginPassword"))
            {
                throw new Exception("Linkedin账号池文件里没有指定Linkedin的用户loginName和密码loginPassword信息");
            }
            string loginName = userInfo["loginName"];
            string loginPassword = userInfo["loginPassword"];
            return Login(runPage, loginName, loginPassword, loginPageUrl, loginSucceedCheckUrl);
        }
        public static bool Login(IRunWebPage runPage, string loginName, string loginPassword, string loginPageUrl, string loginSucceedCheckUrl)
        {
            try
            {
                if (loginName == null || loginName.Length == 0 || loginPassword == null || loginPassword.Length == 0)
                {
                    throw new Exception("用户名loginName和密码loginPassword不可为空");
                }

                ShowLoginPageAndDoLogin(runPage, loginPageUrl, loginSucceedCheckUrl, SysConfig.WebPageRequestTimeout, loginName, loginPassword);
            }
            catch (Exception ex)
            {
                throw new Exception("登录到Linkedin失败!", ex);
            }
            return false;
        }

        private static Dictionary<string, string> GetRowInExcelList(IRunWebPage runPage)
        {
            string excelFilePath = runPage.ExcelFilePath;
            ExcelReader er = new ExcelReader(excelFilePath, "List"); 
            return er.GetFieldValues(0); 
        }

        private static void ShowLoginPageAndDoLogin(IRunWebPage runPage, string loginPageUrl, string loginSucceedCheckUrl, int timeout, string loginName, string loginPassword)
        {
            WebBrowser wb = runPage.ShowWebPage(loginPageUrl, "login", timeout, false);
            DoLoginMethod(runPage, wb, loginName, loginPassword);
            runPage.CheckWebBrowserUrl(wb, loginSucceedCheckUrl, false, timeout);
            runPage.CloseWebPage("login");
        }

        public static void DoLoginMethod(IRunWebPage runPage, WebBrowser wb, string loginName, string loginPassword)
        {
            string scriptMethodCode = "function myDoLogin(loginName, loginPassword){"
                + "document.getElementById('session_key-login').value = loginName;"
                + "document.getElementById('session_password-login').value = loginPassword;"
                + "document.getElementById('btn-primary').click();"
                + "}";
            runPage.InvokeAddScriptMethod(wb, scriptMethodCode, null);
            runPage.InvokeDoScriptMethod(wb, "myDoLogin", new object[] { loginName, loginPassword });
        }

        public static void ShowLogoutPageAndDoLogout(IRunWebPage runPage, string logoutPageUrl, string logoutSucceedCheckUrl, int timeout)
        {
            WebBrowser wb = runPage.ShowWebPage(logoutPageUrl, "logout", timeout, false);
            DoLogoutMethod(runPage, wb, logoutSucceedCheckUrl, timeout);
            runPage.CloseWebPage("logout");
        }

        private static void DoLogoutMethod(IRunWebPage runPage, WebBrowser wb, string logoutSucceedCheckUrl, int timeout)
        {
            string scriptMethodCode = "function myGetLogoutPageUrl(){"
                + "var logoutElements = $('.account-submenu-split-link');"
                + "return (logoutElements.length == 0) ? 'http://www.linkedin.com/logout' : $(logoutElements[0]).attr('href');"
                + "}";
            runPage.InvokeAddScriptMethod(wb, scriptMethodCode, null);
            string logoutPageUrl = (string)runPage.InvokeDoScriptMethod(wb, "myGetLogoutPageUrl", null);
            if (logoutPageUrl != null && logoutPageUrl.Length != 0)
            {
                runPage.ShowWebPage(logoutPageUrl, "logout", timeout, false);
            }
        }

        #region 原来使用的方法
        /*
        public static bool Login(IRunWebPage runPage, string logoutPageUrl, string logoutSucceedCheckUrl, string loginPageUrl, string loginSucceedCheckUrl)
        {
            try
            {
                Dictionary<string, string> row = GetRowInExcelList(runPage);
                if (!row.ContainsKey("loginName") || !row.ContainsKey("loginPassword"))
                {
                    throw new Exception("导入的任务excel中，没有包含用户名loginName和密码loginPassword列");
                }
                if (!row.ContainsKey("keyWords"))
                {
                    throw new Exception("导入的任务excel中，没有包含关键字keyWords列");
                }

                string loginName = row["loginName"];
                string loginPassword = row["loginPassword"];
                string keyWords = row["keyWords"];

                ShowLogoutPageAndDoLogout(runPage, logoutPageUrl, logoutSucceedCheckUrl, SysConfig.WebPageRequestTimeout);
                ShowLoginPageAndDoLogin(runPage, loginPageUrl, loginSucceedCheckUrl, SysConfig.WebPageRequestTimeout, loginName, loginPassword);

            }
            catch (Exception ex)
            {
                throw new Exception("登录到Linkedin失败!", ex);
            }
            return false;
        }

        private static Dictionary<string, string> GetRowInExcelList(IRunWebPage runPage)
        {
            string excelFilePath = runPage.ExcelFilePath;
            ExcelReader er = new ExcelReader(excelFilePath, "List");
            int rowCount = er.GetRowCount();
            if (rowCount > 1)
            {
                throw new Exception("导入的任务excel文件最多包含一条记录");
            }
            else if (rowCount == 0)
            {
                throw new Exception("导入的任务excel文件必须包含一条记录");
            }
            else
            {
                return er.GetFieldValues(0);
            }
        }

        private static void ShowLogoutPageAndDoLogout(IRunWebPage runPage, string logoutPageUrl, string logoutSucceedCheckUrl, int timeout)
        {
            WebBrowser wb = runPage.ShowWebPage(logoutPageUrl, "home", timeout, false);
            DoLogoutMethod(runPage, wb, logoutSucceedCheckUrl, timeout);
        }

        private static void DoLogoutMethod(IRunWebPage runPage, WebBrowser wb, string logoutSucceedCheckUrl, int timeout)
        {
            string scriptMethodCode = "function myGetLogoutPageUrl(){"
                + "var logoutElements = $('.account-submenu-split-link');"
                + "return (logoutElements.length == 0) ? '' : $(logoutElements[0]).attr('href');"
                + "}";
            runPage.InvokeAddScriptMethod(wb, scriptMethodCode, null);
            string logoutPageUrl = (string)runPage.InvokeDoScriptMethod(wb, "myGetLogoutPageUrl", null);
            if (logoutPageUrl != null && logoutPageUrl.Length != 0)
            {
                runPage.ShowWebPage(logoutPageUrl, "home", timeout, false);
            }
        }
        
        private static void ShowLoginPageAndDoLogin(IRunWebPage runPage, string loginPageUrl, string loginSucceedCheckUrl, int timeout, string loginName, string loginPassword)
        {
            WebBrowser wb = runPage.ShowWebPage(loginPageUrl, "home", timeout, false);
            DoLoginMethod(runPage, wb, loginSucceedCheckUrl, loginName, loginPassword, timeout);
        }

        private static void DoLoginMethod(IRunWebPage runPage, WebBrowser wb, string loginSucceedCheckUrl, string loginName, string loginPassword, int timeout)
        {
            string scriptMethodCode = "function myDoLogin(loginName, loginPassword){"
                + "document.getElementById('login-email').value = loginName;"
                + "document.getElementById('login-password').value = loginPassword;"
                + "var parentNode = document.getElementById('login-password').parentNode;"
                + "var allInputChildren = parentNode.getElementsByTagName('input');"
                + "var submitButton = allInputChildren[allInputChildren.length - 1];"
                + "submitButton.click();"
                + "}";
            runPage.InvokeAddScriptMethod(wb, scriptMethodCode, null);
            runPage.InvokeDoScriptMethod(wb, "myDoLogin", new object[] { loginName, loginPassword });
            runPage.CheckWebBrowserUrl(wb, loginSucceedCheckUrl, false, timeout);
        }
        */
        #endregion
    }
}
