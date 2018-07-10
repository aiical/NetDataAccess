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

namespace NetDataAccess.Extended.Jianzhu.QGJZSCJGGGFWPT_RY
{
    public class GetRYQueryUrlList : ExternalRunWebPage
    {          
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            ExcelWriter ew = this.GetExcelWriter();
            for (int i = 0; i < 99; i++)
            {
                string code = i.ToString().PadLeft(2, '0');
                Dictionary<string, string> f2vs = new Dictionary<string, string>(); 
                f2vs.Add("detailPageUrl", "http://jzsc.mohurd.gov.cn/dataservice/query/staff/list?_=" + code);
                f2vs.Add("detailPageName", code);
                f2vs.Add("cookie", "filter_comp=; JSESSIONID=869A8BA542331894C3EE0830A02BB34F; Hm_lvt_b1b4b9ea61b6f1627192160766a9c55c=1515742554,1516079327,1516091394,1516159353; Hm_lpvt_b1b4b9ea61b6f1627192160766a9c55c=1516160300");
                f2vs.Add("formData", "ry_type=&ry_reg_type=&ry_name=&reg_seal_code=&ry_cardno=&ry_qymc=&complexname=" + code);
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

            string resultFilePath = Path.Combine(exportDir, "项目数据_人员列表页首页.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic);
            return resultEW;
        }
    }
}