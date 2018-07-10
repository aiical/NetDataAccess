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

namespace NetDataAccess.Extended.Jianzhu.QGJZSCJGGGFWPT
{
    public class GetQualPageUrlList : ExternalRunWebPage
    {

        private int QualCountPerPage
        {
            get
            {
                return 10;
            }
        }

        #region 构造资质页面url
        public override bool AfterAllGrab(IListSheet listSheet)
        {   
            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{ 
                "detailPageUrl",
                "detailPageName", 
                "cookie",
                "grabStatus", 
                "giveUpGrab",
                "pageIndex",
                "qualCount"});
             
            string exportDir = this.RunPage.GetExportDir();

            string resultFilePath = Path.Combine(exportDir, "企业数据_资质类型列表页.xlsx");


            Dictionary<string, string> resultColumnFormat = new Dictionary<string, string>();
            resultColumnFormat.Add("pageIndex", "#,##0");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, resultColumnFormat); 
             
            GetQuals(listSheet, resultEW);

            //保存到硬盘
            resultEW.SaveToDisk();

            return true;
        }
        #endregion

        #region 构造资质页面url
        /// <summary>
        /// GetCities
        /// </summary>
        /// <param name="listSheet"></param> 
        /// <param name="resultEW"></param>
        private void GetQuals(IListSheet listSheet, ExcelWriter resultEW)
        {
            string[] paramStrs = this.Parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            int qualCount = int.Parse(paramStrs[0]);
            int pageIndex = 0;
            while (pageIndex * this.QualCountPerPage < qualCount)
            {
                Dictionary<string, object> f2vs = new Dictionary<string, object>();
                f2vs.Add("detailPageUrl", "http://jzsc.mohurd.gov.cn/asite/qualapt/aptData?apt_type=&t=" + pageIndex.ToString());
                f2vs.Add("detailPageName", pageIndex.ToString());
                f2vs.Add("cookie", "filter_comp=show; JSESSIONID=DC4BC03F99DEDEBEFEE739B680BC5230; Hm_lvt_b1b4b9ea61b6f1627192160766a9c55c=1513578016,1513646440,1514281557,1514350446; Hm_lpvt_b1b4b9ea61b6f1627192160766a9c55c=1514356771");
                f2vs.Add("pageIndex", pageIndex);
                f2vs.Add("qualCount", qualCount);
                resultEW.AddRow(f2vs);
                pageIndex++;
            }
        }
        #endregion
    }
}