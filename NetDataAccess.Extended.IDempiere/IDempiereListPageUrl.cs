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
using NPOI.XSSF.UserModel;
using NetDataAccess.Base.DB;
using System.Web; 

namespace NetDataAccess.Extended.IDempiere
{
    public class IDempiereListPageUrl : CustomProgramBase
    { 
        public bool Run(string parameters, IListSheet listSheet )
        {
            try
            {
                return this.GetListPageUrl(listSheet, parameters); 
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        private bool GetListPageUrl(IListSheet listSheet, string parameters)
        {
            string[] paramStrs = parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            int toIndex = int.Parse(paramStrs[0]); 
            string exportDir = paramStrs[1];
            bool succeed = true;  

            Dictionary<string, int> subjectColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                "detailPageUrl",
                "detailPageName",
                "cookie",
                "grabStatus",
                "giveUpGrab",
                "fromIndex",
                "toIndex"});
            string subjectFileExcelPath = Path.Combine(exportDir, this.RunPage.Project.Name + "_List.xlsx");
            ExcelWriter subjectEW = new ExcelWriter(subjectFileExcelPath, "List", subjectColumnDic);

            int i = 1;
            while (i <= toIndex)
            {
                string fromIndexStr = i.ToString();
                string toIndexStr = (i + 19).ToString();
                string url = "https://groups.google.com/forum/?_escaped_fragment_=forum/idempiere%5B" + fromIndexStr + "-" + toIndexStr + "%5D";
                Dictionary<string, string> row = new Dictionary<string, string>();
                row.Add("detailPageUrl", url);
                row.Add("detailPageName", fromIndexStr + "_" + toIndexStr);
                row.Add("fromIndex", fromIndexStr);
                row.Add("toIndex", toIndexStr);
                subjectEW.AddRow(row);
                i = i + 20;
            } 
            subjectEW.SaveToDisk();
            return succeed;
        }
    }
}