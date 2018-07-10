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
using System.Web;
using NetDataAccess.Base.Reader;

namespace NetDataAccess.Extended.Jianzhu.QGJZSCJGGGFWPT_XM
{
    public class GetProjectSummary : ExternalRunWebPage
    { 
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            try
            {
                return this.GetSummary(listSheet);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private ExcelWriter GetSummaryExcelWriter(string exportDir)
        { 
            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("pCode", 0);
            resultColumnDic.Add("count", 1); 

            string resultFilePath = Path.Combine(exportDir, "项目数据_统计信息"+DateTime.Now.ToString("yyyyMMddHHmmss")+".xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }

        private bool GetSummary(IListSheet listSheet)
        {
            string[] parameters = this.Parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            string projectListDir = parameters[0];
            string listFileName = parameters[1];
            int listFileCount = int.Parse(parameters[2]);
            string exportDir = parameters[3];
            Dictionary<string, int> summaryDic = new Dictionary<string, int>();
            for (int i = 0; i < listFileCount; i++)
            {
                string filePath = Path.Combine(projectListDir, listFileName + "_" + (i + 1).ToString() + ".xlsx");
                ExcelReader er = new ExcelReader(filePath, "List");
                int rowCount = er.GetRowCount();
                for (int j = 0; j < rowCount; j++)
                {
                    Dictionary<string, string> row = er.GetFieldValues(j);
                    string code = row["detailPageName"];
                    string pCode = code.Substring(0, 6);
                    if (!summaryDic.ContainsKey(pCode))
                    {
                        summaryDic.Add(pCode, 0);
                    }
                    summaryDic[pCode] = summaryDic[pCode] + 1;
                }
            }
            ExcelWriter ew = this.GetSummaryExcelWriter(exportDir);
            foreach (string pCode in summaryDic.Keys)
            {
                Dictionary<string, string> f2vs = new Dictionary<string, string>();
                f2vs.Add("pCode", pCode);
                f2vs.Add("count", summaryDic[pCode].ToString());
                ew.AddRow(f2vs);
            }

            ew.SaveToDisk();
            return true;
        }
    }
}