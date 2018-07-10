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

namespace NetDataAccess.Extended.Jianzhu.JinanLouPan
{
    public class GetLoupanPageUrlList : ExternalRunWebPage
    {

        private int QualCountPerPage
        {
            get
            {
                return 10;
            }
        } 
        public override bool AfterAllGrab(IListSheet listSheet)
        {   
            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{ 
                "detailPageUrl",
                "detailPageName", 
                "cookie",
                "grabStatus", 
                "giveUpGrab",
                "pageIndex"});
             
            string exportDir = this.RunPage.GetExportDir();

            string resultFilePath = Path.Combine(exportDir, "济南楼盘_楼盘列表页.xlsx");


            Dictionary<string, string> resultColumnFormat = new Dictionary<string, string>();
            resultColumnFormat.Add("pageIndex", "#,##0");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, resultColumnFormat); 
             
            GetQuals(listSheet, resultEW);

            //保存到硬盘
            resultEW.SaveToDisk();

            return true;
        }  
        /// <summary>
        /// GetCities
        /// </summary>
        /// <param name="listSheet"></param> 
        /// <param name="resultEW"></param>
        private void GetQuals(IListSheet listSheet, ExcelWriter resultEW)
        {
            string[] paramStrs = this.Parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            int buildingCount = int.Parse(paramStrs[0]);
            int pageIndex = 1;
            while (pageIndex <= buildingCount)
            {
                Dictionary<string, object> f2vs = new Dictionary<string, object>();
                f2vs.Add("detailPageUrl", "http://www.jnfdc.gov.cn/onsaling/index_" + pageIndex.ToString() + ".shtml");
                f2vs.Add("detailPageName", pageIndex.ToString());
                f2vs.Add("pageIndex", pageIndex); 
                resultEW.AddRow(f2vs);
                pageIndex++;
            }
        } 
    }
}