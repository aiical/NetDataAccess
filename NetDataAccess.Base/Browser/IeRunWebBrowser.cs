using mshtml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NetDataAccess.Base.Browser
{
    public class IeRunWebBrowser : WebBrowser,IWebBrowser
    {
        public IeRunWebBrowser()
        {
            this.DocumentCompleted += IeRunWebBrowser_DocumentCompleted; 
        } 
        private void IeRunWebBrowser_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            if (DocumentLoadCompleted != null)
            {
                DocumentLoadCompleted(this);
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

        public bool Loaded()
        {
            return true;//return !this.IsBusy && this.ReadyState == WebBrowserReadyState.Complete;
        }

        public string GetDocumentText()
        {
            return this.Document == null ? null : (this.Document.Body == null ? null : this.Document.Body.OuterHtml);
        }
        public void AvoidWebBrowserUnfriendlyJavaScript()
        {
            HtmlElement sElement = this.Document.CreateElement("script");
            IHTMLScriptElement scriptElement = (IHTMLScriptElement)sElement.DomElement;
            scriptElement.text = "alert = function(){};confirm = function(){};";
            this.Document.Body.AppendChild(sElement);
        }
        public void ScrollTo(int x, int y)
        { 
            this.Document.Window.ScrollTo(x, y); 
        }

        private int GetWebPageHeight()
        {
            return this.Document.Body.OffsetRectangle.Bottom;
        }
        public IWebBrowserDocumentCompletedEventHandler DocumentLoadCompleted { get; set; }


        public void SetControlValueById(string id, string attributeName, string attributeValue)
        { 
            HtmlElement element = this.Document.GetElementById(id);
            element.SetAttribute(attributeName, attributeValue);
        }
        public void AddScriptMethod(string scriptMethodCode)
        { 
            HtmlElement sElement = this.Document.CreateElement("script");
            IHTMLScriptElement scriptElement = (IHTMLScriptElement)sElement.DomElement;
            scriptElement.text = scriptMethodCode;
            this.Document.Body.AppendChild(sElement);
        }

        public object DoScriptMethod(string methodName, object[] parameters)
        {
            return this.Document.InvokeScript(methodName, parameters);
        }

        public string GetWebBrowserPageUrlMethod()
        {
            return this.Url.AbsoluteUri;
        }
    }
}
