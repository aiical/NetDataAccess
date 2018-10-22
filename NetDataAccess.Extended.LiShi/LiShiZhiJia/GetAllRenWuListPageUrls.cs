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
using NetDataAccess.Base.Reader;

namespace NetDataAccess.Extended.LiShi.LiShiZhiJia
{
    public class GetAllRenWuListPageUrls : ExternalRunWebPage
    { 
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            this.GetListPageUrls(listSheet); 
            return true;
        }

        private void GetListPageUrls(IListSheet listSheet)
        {
            ExcelWriter ew = this.CreateWriter();

            for (int i = 1; i <= 17; i++)
            {
                string url = "http://www.lszj.com/renwu/4e2d56fd-all-ALL-"+i.ToString()+".html";
                Dictionary<string, string> row = new Dictionary<string, string>();
                row.Add("detailPageUrl", url);
                row.Add("detailPageName", i.ToString());  
                ew.AddRow(row);
            }
            ew.SaveToDisk();
        }

        private ExcelWriter CreateWriter()
        {
            String exportDir = this.RunPage.GetExportDir(); 
            string resultFilePath = Path.Combine(exportDir, "历史_历史之家_人物列表页.xlsx");

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }
         
    }
}