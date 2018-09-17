using mshtml;
using NetDataAccess.Base.Config;
using NetDataAccess.Base.EnumTypes;
using NetDataAccess.Base.UI;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using HtmlAgilityPack;
using System.Windows.Forms;
using mshtml;
using System.IO;
using NetDataAccess.Base.Server;

namespace NetDataAccess.Extended.Luyou
{
    public class InputAnjukeIdCode
    {


        public InputAnjukeIdCode(IRunWebPage runPage)
        {
            this._RunPage = runPage;
        }

        private IRunWebPage _RunPage = null;
        protected IRunWebPage RunPage
        {
            get
            {
                return this._RunPage;
            }
        }
        private string _PageUrl = "https://www.anjuke.com/captcha-verify/?callback=shield&from=antispam&history=aHR0cHM6Ly9tLmFuanVrZS5jb20vam4vc2FsZS9BMTExNjAwOTk1OC8%2FaXNhdWN0aW9uPTIwMSZwb3NpdGlvbj01NDQma3d0eXBlPWZpbHRlcg%3D%3D";
        private string PageUrl
        {
            get
            {
                return this._PageUrl;
            }  
        }

        public void BeginProcessIdCode(string tabName)
        {

            //允许跳转到查询页面的次数，有时会出现跳转至登录页面的情况
            const int allowGoToQueryPageCount = 10;

            int goToQueryPageErrorCount = 0;

            string currentUrl = "";
            WebBrowser webBrowser = null;

            while (currentUrl != this.PageUrl && goToQueryPageErrorCount < allowGoToQueryPageCount)
            {
                //加载网页
                webBrowser = this.ShowWebPage(this.PageUrl, tabName);

                currentUrl = webBrowser.Url.ToString();
            }

            if (currentUrl != _PageUrl)
            {
                throw new Exception("无法打开页面.");
            }
            else
            {

                this.RunPage.InvokeAppendLogText("准备识别验证码", LogLevelType.Normal, true);
                InvokeProcess(webBrowser);
            }
        }

        #region 开始识别验证码
        private void InvokeProcess(WebBrowser webBrowser)
        {
            string code = this.GetTxtCode(webBrowser);
            Thread.Sleep(1000);
            this.InvokeInputCode(webBrowser, code);
        }

        private void InvokeInputCode(WebBrowser webBrowser, string code)
        {
            //采用异步获取，因为跨线程了
            webBrowser.Invoke(new InputCodeDelegate(InputCode), new object[] { webBrowser, code });
        }
        private delegate void InputCodeDelegate(WebBrowser webBrowser, string code);
        private void InputCode(WebBrowser webBrowser, string code)
        {
            webBrowser.Document.GetElementById("code").SetAttribute("value", code.ToLower());
            HtmlElementCollection inputCtrls = webBrowser.Document.GetElementsByTagName("input");
            foreach (HtmlElement inputCtrl in inputCtrls)
            {
                if (inputCtrl.GetAttribute("value") == "提交")
                {
                    inputCtrl.InvokeMember("click");
                    this.RunPage.InvokeAppendLogText("已输入验证码", LogLevelType.System, true);
                    break;
                }
            } 

        }

        #endregion

        #region 在网页中获取图片
        private Image InvokeGetWebImage(WebBrowser webBrowser, string imgeTagId)
        {
            //采用异步获取，因为跨线程了
            return (Image)webBrowser.Invoke(new GetWebImageInvokeDelegate(GetWebImage), new object[] { webBrowser, imgeTagId });
        }
        private delegate Image GetWebImageInvokeDelegate(WebBrowser webCtl, string imgeTagId);
        /// <summary>
        /// 返回指定WebBrowser中图片<IMG></IMG>中的图内容
        /// </summary>
        /// <param name="webCtl">WebBrowser控件</param>
        /// <param name="imgeTag">IMG元素</param>
        /// <returns>IMG对象</returns>
        private Image GetWebImage(WebBrowser webBrowser, string imgeTagId)
        {
            System.Windows.Forms.HtmlDocument winHtmlDoc = webBrowser.Document;
            HTMLDocument doc = (HTMLDocument)winHtmlDoc.DomDocument;
            HtmlElement imgTag = winHtmlDoc.GetElementById(imgeTagId);
            HTMLBody body = (HTMLBody)doc.body;
            IHTMLControlRange rang = (IHTMLControlRange)body.createControlRange();
            IHTMLControlElement imgE = (IHTMLControlElement)imgTag.DomElement; //图片地址
            rang.add(imgE);
            rang.execCommand("Copy", false, null);  //拷贝到内存
            Image numImage = Clipboard.GetImage();
            return numImage;
        }
        #endregion

        #region 获取验证码并识别
        public string GetTxtCode(WebBrowser webBrowser)
        { 
            Image souceImg = null;
            try
            {
                //获取验证码图片
                souceImg = InvokeGetWebImage(webBrowser, "seccode");
                string imagePath = Path.Combine(Application.StartupPath, "anjukecode.jpg");
                souceImg.Save(imagePath);

                StringBuilder codeStr = new StringBuilder();
                int result = Dama.D2File("e34773ffd51e5e7f1f1b6d7ee4b24c29", "moker1018", "111111", imagePath, 30, 92056, codeStr); 
                if (result >= 0)
                {
                    return codeStr.ToString();
                }
                else
                {
                    throw new Exception("识别失败, resultCode = " + result.ToString());
                }

            }
            catch (Exception ex)
            {
                this.RunPage.InvokeAppendLogText(ex.Message, LogLevelType.Error, true);
                throw ex;
            }
            finally
            {
                if (souceImg != null)
                {
                    souceImg.Dispose();
                } 
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
        private WebBrowser ShowWebPage(string url, string tabName)
        {
            WebBrowser webBrowser = (WebBrowser)this.RunPage.InvokeShowWebPage(url, tabName, WebBrowserType.IE);
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

            //再增加个等待，等待异步加载的数据
            Thread.Sleep(1000);
            return webBrowser;
        }
        #endregion
    }
}
