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
    public class GetProjectQueryUrlList : ExternalRunWebPage
    {          
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            string[] paramParts = this.Parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            string companyProjectListFilePath = paramParts[0];

            ExcelReader er = new ExcelReader(companyProjectListFilePath,"List");
            int rowCount = er.GetRowCount();
            Dictionary<string, string> codeDic = new Dictionary<string, string>();
            List<string> codeList = new List<string>();
            for (int i = 0; i < rowCount; i++)
            {
                Dictionary<string,string> row = er.GetFieldValues(i);
                string code = row["code"];
                if (!codeDic.ContainsKey(code))
                {
                    codeDic.Add(code, null);
                    codeList.Add(code);
                }
            }
             
            int codeCount = codeList.Count;
            ExcelWriter ew = this.GetExcelWriter();
            for (int i = 0; i < codeCount; i++)
            { 
                Dictionary<string, string> f2vs = new Dictionary<string, string>();
                string code = codeList[i];
                f2vs.Add("detailPageUrl", "http://jzsc.mohurd.gov.cn/dataservice/query/project/list?_=" + code);
                f2vs.Add("detailPageName", code);
                f2vs.Add("cookie", "filter_comp=; JSESSIONID=F1DC2E6DC10B3E64CC59C070A5722639; Hm_lvt_b1b4b9ea61b6f1627192160766a9c55c=1515293273,1515384893,1515553333,1515638274; Hm_lpvt_b1b4b9ea61b6f1627192160766a9c55c=1515645794");
                f2vs.Add("formData", "jsxm_name=&cons_name=&jsxm_region=&jsxm_region_id=&complexname=" + code);
                ew.AddRow(f2vs);
            }
            ew.SaveToDisk();

            return true;
        }
        private ExcelWriter GetExcelWriter()
        {
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("formData", 5); 

            string resultFilePath = Path.Combine(exportDir, "项目数据_项目列表页首页.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic);
            return resultEW;
        }
    }
}