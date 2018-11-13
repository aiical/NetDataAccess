using CefSharp;
using NetDataAccess.Base.Browser;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NetDataAccess.Extended.ShangBiao.Common
{
    public class LifeSpanHandler : ILifeSpanHandler
    {
        private ChromiumRunWebBrowser _PopBrowser = null;
        public ChromiumRunWebBrowser PopBrowser
        {
            get
            {
                return this._PopBrowser;
            }
            set
            {
                this._PopBrowser = value;
            }
        }

        public bool DoClose(CefSharp.IWebBrowser browserControl, IBrowser browser)
        {
            return true;
        }
        public void OnAfterCreated(CefSharp.IWebBrowser browserControl, IBrowser browser)
        {
        }
        public void OnBeforeClose(CefSharp.IWebBrowser browserControl, IBrowser browser)
        {
        }
        public bool OnBeforePopup(CefSharp.IWebBrowser browserControl, IBrowser browser, IFrame frame, string targetUrl, string targetFrameName, WindowOpenDisposition targetDisposition, bool userGesture, IWindowInfo windowInfo, ref bool noJavascriptAccess, out CefSharp.IWebBrowser newBrowser)
        {
            newBrowser = null;
            return true;
        }
    }
}
