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

namespace NetDataAccess.Extended.Jianzhu.QGJZSCJGGGFWPT
{
    public class GetListPageUrlByCertno : ExternalRunWebPage
    {

        #region GetProvinces
        public override bool AfterAllGrab(IListSheet listSheet)
        { 
            //已经下载下来的首页html保存到的目录（文件夹）
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            
            //输出excel表格包含的列
            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{ 
                "detailPageUrl",
                "detailPageName", 
                "cookie",
                "grabStatus", 
                "giveUpGrab", 
                "certNo"});
             
            string exportDir = this.RunPage.GetExportDir();

            string resultFilePath = Path.Combine(exportDir, "企业数据_证书编码查询企业列表页首页.xlsx");
             
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic);
             
            GetListPageUrls(listSheet, pageSourceDir, resultEW);
             
            resultEW.SaveToDisk();

            return true;
        }
        #endregion

        #region GetListPageUrls
        /// <summary>
        /// GetListPageUrls
        /// </summary>
        /// <param name="listSheet"></param>
        /// <param name="pageSourceDir"></param>
        /// <param name="resultEW"></param>
        private void GetListPageUrls(IListSheet listSheet, string pageSourceDir, ExcelWriter resultEW)
        {
            for (int i = 0; i < 10000; i++)
            {
                string certNo = i.ToString().PadLeft(4, '0'); 
                Dictionary<string, string> f2vs = new Dictionary<string, string>();
                f2vs.Add("detailPageUrl", "http://jzsc.mohurd.gov.cn/dataservice/query/comp/list?certNo=" + certNo);
                f2vs.Add("detailPageName", certNo);
                f2vs.Add("cookie", "filter_comp=show; JSESSIONID=DC4BC03F99DEDEBEFEE739B680BC5230; Hm_lvt_b1b4b9ea61b6f1627192160766a9c55c=1513578016,1513646440,1514281557,1514350446; Hm_lpvt_b1b4b9ea61b6f1627192160766a9c55c=1514356771");
                f2vs.Add("certNo", certNo); 
                resultEW.AddRow(f2vs);
            }
        }
        #endregion
    }
}