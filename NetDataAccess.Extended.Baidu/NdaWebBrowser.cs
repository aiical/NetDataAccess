using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NetDataAccess.Extended.Baidu
{
    public class NdaWebBrowser: WebBrowser
    {
        private string _TabName = "";
        public string TabName
        {
            get;
            set;
        }
    }
}
