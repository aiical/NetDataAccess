using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NetDataAccess.Base.Browser
{
    public interface IWebBrowser
    { 
        bool Loaded();

        string TabName { get; set; }

        string GetDocumentText();

        void Navigate(string url);

        object Invoke(Delegate method, params object[] args); 

        IWebBrowserDocumentCompletedEventHandler DocumentLoadCompleted { get; set; }
        
        void Dispose();

        DockStyle Dock { get; set; }

        bool ScriptErrorsSuppressed { get; set; }

        bool AllowWebBrowserDrop { get; set; }

        void AvoidWebBrowserUnfriendlyJavaScript();

        void ScrollTo(int x, int y);

        object ObjectForScripting { get; set; }

        void SetControlValueById(string id, string attributeName, string attributeValue);

        void AddScriptMethod(string scriptMethodCode);

        object DoScriptMethod(string methodName, object[] parameters);

        bool GoBack();

        string GetWebBrowserPageUrlMethod();
    }
    public delegate void IWebBrowserDocumentCompletedEventHandler(IWebBrowser webBrowser);
}
