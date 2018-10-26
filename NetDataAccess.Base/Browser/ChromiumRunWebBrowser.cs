using CefSharp;
using CefSharp.WinForms;
using mshtml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NetDataAccess.Base.Browser
{
    public class ChromiumRunWebBrowser : ChromiumWebBrowser,IWebBrowser
    {
        public ChromiumRunWebBrowser():
            base("about:blank")
        //base("http://www.baidu.com")
        {
            this.FrameLoadEnd += ChromiumRunWebBrowser_FrameLoadEnd; 
        } 
        private bool _FrameLoaded = false;

        public new void Load(string url)
        { 
            base.Load(url);
        }

        private void ChromiumRunWebBrowser_FrameLoadEnd(object sender, CefSharp.FrameLoadEndEventArgs e)
        {
            if (this.Address == "about:blank")
            {
                if (this._TargetUrl.Length != 0)
                {
                    this._BlankLoaded = true;
                    this.Load(this._TargetUrl);
                }
            }
            else{
                _FrameLoaded = true;
                if (DocumentLoadCompleted != null)
                {
                    DocumentLoadCompleted(this);
                }
            }
        }

        private string _TabName = "";
        public string TabName
        {
            get
            {
                return this._TabName;
            }
            set
            {
                this._TabName = value;
            }
        } 

        public bool IsBusy
        {
            get
            {
                return !_FrameLoaded;
            }
        }

        public bool Loaded()
        {
            return _FrameLoaded;
        }
        private bool _BlankLoaded = false;
        private string _TargetUrl = "";

        public void Navigate(string url)
        {
            _TargetUrl = url;
            if (this._BlankLoaded)
            {
                this.Load(url);
            }
        }

        public IWebBrowserDocumentCompletedEventHandler DocumentLoadCompleted { get; set; }

        public bool ScriptErrorsSuppressed
        { 
            set
            {  
            }
            get
            {
                return false;
            }
        }

        public bool AllowWebBrowserDrop
        {
            set
            {
            }
            get
            {
                return false;
            }
        }
        public string GetDocumentText()
        {
            if (!this.Loaded())
            {
                return null;
            }
            else
            {
                Task<String> taskHtml = this.GetBrowser().MainFrame.GetSourceAsync();
                if (taskHtml.Wait(3000))
                {
                    string response = taskHtml.Result;
                    return response;
                }
                else
                {
                    return null;
                }
            }
        }
        public bool GoBack()
        {
            if (this.CanGoBack)
            {
                this.GetBrowser().GoBack();
                return true;
            }
            else
            {
                return false;
            }
        }

        public void AvoidWebBrowserUnfriendlyJavaScript()
        {
            string code = "alert = function(){};confirm = function(){};";
            this.GetBrowser().MainFrame.ExecuteJavaScriptAsync(code);  
        }

        public void ScrollTo(int x, int y)
        {
            string code = "window.scrollTo(" + x.ToString() + "," + y.ToString() + ");";
            this.GetBrowser().MainFrame.ExecuteJavaScriptAsync(code);   
        }

        private object _ObjectForScripting = null;
        public object ObjectForScripting
        {
            get
            {
                return this._ObjectForScripting;
            }
            set
            {
                this.RegisterAsyncJsObject("callbackObject", value);
                this._ObjectForScripting = value;
            }
        }

        public void SetControlValueById(string id, string attributeName, string attributeValue)
        {
            string code = "docuemnt.getElementById(\"" + id + "\")." + attributeName + " = \"" + attributeValue + "\";";
            this.GetBrowser().MainFrame.ExecuteJavaScriptAsync(code);  
        }

        public void AddScriptMethod(string scriptMethodCode)
        {
            this.GetBrowser().MainFrame.ExecuteJavaScriptAsync(scriptMethodCode);
        }


        public object DoScriptMethod(string methodName, object[] parameters)
        {
            StringBuilder code = new StringBuilder(methodName + "(");
            if (parameters != null)
            {
                for (int i = 0; i < parameters.Length; i++)
                {
                    if (i > 0)
                    {
                        code.Append(",");
                    }
                    string p = parameters[i] == null ? "" : parameters[i].ToString().Replace("\"", "\\\"");
                    code.Append("\"" + p + "\"");
                }
            }
            code.Append(")");


            Task<JavascriptResponse> response = this.GetBrowser().MainFrame.EvaluateScriptAsync(code.ToString());
            response.Wait();
            return response.Result.Result;
        }

        public string GetWebBrowserPageUrlMethod()
        {
            return this.GetBrowser().MainFrame.Url;
        }
    }
}
