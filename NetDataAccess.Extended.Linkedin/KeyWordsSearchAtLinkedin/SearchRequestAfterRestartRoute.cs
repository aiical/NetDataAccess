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
using NetDataAccess.Base.Writer;
using NetDataAccess.Base.Reader;
using NetDataAccess.Base.CsvHelper;
using NetDataAccess.Base.DB;
using HtmlAgilityPack;
using System.Web;
using System.Runtime.Remoting;
using System.Reflection;
using System.Collections;
using NetDataAccess.Base.Browser;

namespace NetDataAccess.Extended.Linkedin.KeyWordsSearchAtLinkedin
{
    [System.Runtime.InteropServices.ComVisibleAttribute(true)]
    public class SearchRequestAfterRestartRoute : ExternalRunWebPage
    {
        #region _Succeed 
        private bool _RestartSucceed = false;
        private string _PageUrl = "";
        private string _User = "";
        private string _Password = "";
        private int _RestartTimeout = 0; 
        private int _BatchReqeustCount = 5; 
        #endregion

        public override void CheckRequestCompleteFile(string webPageText, Dictionary<string, string> listRow)
        {
            HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
            htmlDoc.LoadHtml(webPageText);

            HtmlNode inputElement =  htmlDoc.DocumentNode.SelectSingleNode("//input[@id=\"lastName\"]");
            if (inputElement == null)
            {
                throw new Exception("返回的文本页面不准确.");
            } 
        }
        public override bool AfterGrabOneCatchException(string pageUrl, Dictionary<string, string> listRow, Exception ex)
        {
            BeginRestart();
            return base.AfterGrabOneCatchException(pageUrl, listRow, ex);
        }

        public override void WebRequestHtml_BeforeSendRequest(string pageUrl, Dictionary<string, string> listRow, Base.Web.NDAWebClient client)
        {
            string userAgent = this.RunPage.CurrentUserAgents.GetOnePcUserAgent();
            client.Headers["User-Agent"] = userAgent; 
        }

        #region Run
        public override bool BeforeAllGrab()
        {
            this.GetParameters();
            return true;
        }
        private void GetParameters()
        {
            string[] parameterArray = this.Parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            _PageUrl = parameterArray[0];
            _User = parameterArray[1];
            _Password = parameterArray[2];
            _RestartTimeout = int.Parse(parameterArray[3]); 
            _BatchReqeustCount = int.Parse(parameterArray[4]); 
        }
        #endregion 

        private int _RequestCount = 0;

        /*
        public override void AfterGrabOne(string pageUrl, Dictionary<string, string> listRow, bool needReGrab, bool existLocalFile)
        {
            _RequestCount++;
            if (_RequestCount == _BatchReqeustCount)
            {
                BeginRestart();
                _RequestCount = 0;
            }
        }
         * */
         
        private int _GrabPerCheckSecond = 2;

        #region 开始搜索
        private void BeginRestart()
        {
           AutoRestartLuyou();
             //AutoInputAnjukeIdCode();
        }
        #endregion
        
        #region 开始搜索
        private void AutoInputAnjukeIdCode()
        {
            //InputAnjukeIdCode anjukeCode = new InputAnjukeIdCode(this.RunPage);
            //anjukeCode.BeginProcessIdCode("自动输入安居客验证码");
        }
        #endregion
        
        #region 开始搜索
        private void AutoRestartLuyou()
        {
            string tabName = "重启路由"; 
            while (1==1)
            {
                _RestartSucceed = false;
                BeginRestartOnce(tabName);
                int waitingSeconds = 0;
                while (!_RestartSucceed && waitingSeconds < _RestartTimeout)
                {
                    waitingSeconds = waitingSeconds + _GrabPerCheckSecond;
                    Thread.Sleep(_GrabPerCheckSecond * 1000);
                }
                if (waitingSeconds >= _RestartTimeout)
                {
                    //超时 
                    this.RunPage.InvokeAppendLogText("抓取超时", LogLevelType.Error, true); 
                }
                else
                {
                    this.RunPage.InvokeAppendLogText("重新拨号成功.", LogLevelType.System, true);
                }

                Thread.Sleep(15000);
            } 
        }
        private void BeginRestartOnce(string tabName)
        {

            //允许跳转到查询页面的次数，有时会出现跳转至登录页面的情况
            const int allowGoToQueryPageCount = 10;

            int goToQueryPageErrorCount = 0;

            string currentUrl = "";
            IeRunWebBrowser webBrowser = null;

            while (currentUrl != _PageUrl && goToQueryPageErrorCount < allowGoToQueryPageCount)
            {
                //加载网页
                webBrowser = this.ShowWebPage(_PageUrl, tabName); 

                currentUrl = webBrowser.Url.ToString();
            }

            if (currentUrl != _PageUrl)
            {
                throw new Exception("无法打开页面.");
            }
            else
            { 

                this.RunPage.InvokeAppendLogText("准备重新拨号", LogLevelType.Normal, true);
                InvokeRestartNow(webBrowser, this._User, this._Password, tabName);
            }
        }
        #endregion 

        #region 获取网页信息超时时间
        /// <summary>
        /// 获取网页信息超时时间
        /// </summary>
        private int WebRequestTimeout = 20 * 1000;
        #endregion
        
        #region 显示网页
        private IeRunWebBrowser ShowWebPage(string url, string tabName)
        {
            IeRunWebBrowser webBrowser = (IeRunWebBrowser)this.RunPage.InvokeShowWebPage(url, tabName, WebBrowserType.IE); 
            int waitCount = 0;
            while (!this.RunPage.CheckIsComplete(tabName))
            {
                if (SysConfig.WebPageRequestInterval * waitCount > WebRequestTimeout)
                {
                    string errorInfo = "打开页面超时! PageUrl = " + url + ". 但是继续执行!";
                    this.RunPage.InvokeAppendLogText(errorInfo, Base.EnumTypes.LogLevelType.System, true);
                    break;
                    //超时
                    //throw new Exception("打开页面超时. PageUrl = " + url);
                }
                //等待
                waitCount++;
                Thread.Sleep(SysConfig.WebPageRequestInterval);
            }


            this.InvokeAddMyScript(webBrowser);

            //再增加个等待，等待异步加载的数据
            Thread.Sleep(1000);
            return webBrowser;
        }
        #endregion

        #region AddMyScript
        private void InvokeAddMyScript(IeRunWebBrowser webBrowser)
        {
            webBrowser.Invoke(new AddMyScriptInvokeDelegate(AddMyScript), new object[] { webBrowser, "" });
        }
        private delegate void AddMyScriptInvokeDelegate(IeRunWebBrowser webBrowser, string p1);
        private void AddMyScript(IeRunWebBrowser webBrowser, string p1)
        {
            webBrowser.ObjectForScripting = this;
        }
        #endregion 

        #region 开始搜索
        private void InvokeRestartNow(IeRunWebBrowser webBrowser, string user, string password, string tabName)
        {
            webBrowser.Invoke(new LoginInvokeDelegate(Login), new object[] { webBrowser, user, password });
            Thread.Sleep(10000);
            webBrowser.Invoke(new GoonInvokeDelegate(Goon), new object[] { webBrowser });
            Thread.Sleep(10000);
            webBrowser.Invoke(new GoToMaintenancePageInvokeDelegate(GoToMaintenancePage), new object[] { webBrowser });
            Thread.Sleep(10000);
            webBrowser.Invoke(new GoToManagmentPageInvokeDelegate(GoToManagmentPage), new object[] { webBrowser });
            Thread.Sleep(10000);
            webBrowser.Invoke(new GoToRebootPageInvokeDelegate(GoToReboottPage), new object[] { webBrowser });
            Thread.Sleep(10000);
            webBrowser.Invoke(new RebootInvokeDelegate(Reboot), new object[] { webBrowser });
        }

        private delegate void LoginInvokeDelegate(IeRunWebBrowser webBrowser, string user, string password);
        private void Login(IeRunWebBrowser webBrowser, string user, string password)
        {
            webBrowser.Document.GetElementById("txt_usr_name").SetAttribute("value", user);
            webBrowser.Document.GetElementById("txt_password").SetAttribute("value", password);
            webBrowser.Document.GetElementById("btn_logon").InvokeMember("onclick");
        }

        private delegate void GoonInvokeDelegate(IeRunWebBrowser webBrowser);
        private void Goon(IeRunWebBrowser webBrowser)
        {
            HtmlElement btnElement = webBrowser.Document.GetElementById("btn_confirm");
            if (btnElement != null && btnElement.GetAttribute("value") == "继续")
            {
                btnElement.InvokeMember("onclick");
            }
        }

        private delegate void GoToMaintenancePageInvokeDelegate(IeRunWebBrowser webBrowser);
        private void GoToMaintenancePage(IeRunWebBrowser webBrowser)
        {
            HtmlElement btnElement = webBrowser.Document.GetElementById("fst_Maintenance");
            if (btnElement != null)
            {
                btnElement.InvokeMember("onclick");
            }
        }

        private delegate void GoToManagmentPageInvokeDelegate(IeRunWebBrowser webBrowser);
        private void GoToManagmentPage(IeRunWebBrowser webBrowser)
        {
            HtmlElement btnElement = webBrowser.Document.GetElementById("sec_Management");
            if (btnElement != null)
            {
                btnElement.InvokeMember("onclick");
            }
        }

        private delegate void GoToRebootPageInvokeDelegate(IeRunWebBrowser webBrowser);
        private void GoToReboottPage(IeRunWebBrowser webBrowser)
        {
            HtmlElement btnElement = webBrowser.Document.GetElementById("tab_Reboot");
            if (btnElement != null)
            {
                btnElement.InvokeMember("onclick");
            }
        }

        private delegate void RebootInvokeDelegate(IeRunWebBrowser webBrowser);
        private void Reboot(IeRunWebBrowser webBrowser)
        {
            webBrowser.Document.Window.Frames["ifm_func_module"].Document.InvokeScript("showRebootPro");
            _RestartSucceed = true;
        }


        #endregion

    }
}