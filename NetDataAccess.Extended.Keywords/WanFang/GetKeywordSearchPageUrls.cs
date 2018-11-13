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

namespace NetDataAccess.Extended.Keywords.WanFang
{
    public class GetKeywordSearchPageUrls : ExternalRunWebPage
    { 
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            try 
            {
                this.GetSearchPageUrls(listSheet); 
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return true;
        }

        private void GetSearchPageUrls(IListSheet listSheet)
        {
            String[] parameters = this.Parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            string exportDir = parameters[0];

            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[] { 
                "detailPageUrl", 
                "detailPageName", 
                "cookie", 
                "grabStatus", 
                "giveUpGrab", 
                "keyword", 
                "品类", 
                "词类型" });

            string resultFilePath = Path.Combine(exportDir, "万方期刊_专业关键词_搜索页首页.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"];

                string keyword = row["keyword"];
                string pinLei = row["品类"];
                string ciLeiXing = row["词类型"];
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string keywordEncode = CommonUtil.StringToHexString(keyword, Encoding.UTF8);
                    string detailPageUrl = "http://librarian.wanfangdata.com.cn/SearchResult.aspx?dbhit=wf_qk%3a2466%7cwf_xw%3a188%7cwf_hy%3a87%7cnstl_qk%3a0%7cnstl_hy%3a0&q=%e5%85%b3%e9%94%ae%e8%af%8d%3a(%22" + keywordEncode + "%22)+*+Date%3a2015-2018&db=wf_qk%7cwf_xw%7cwf_hy%7cnstl_qk%7cnstl_hy&p=1";
                    string detailPageName = row["keyword"] + "_" + row["品类"];
                    Dictionary<string, string> f2vs = new Dictionary<string, string>();
                    f2vs.Add("detailPageUrl", detailPageUrl);
                    f2vs.Add("detailPageName", detailPageName); 
                    f2vs.Add("keyword", row["keyword"]);
                    f2vs.Add("品类", row["品类"]);
                    f2vs.Add("词类型", row["词类型"]);
                    resultEW.AddRow(f2vs);
                }
            }
            resultEW.SaveToDisk();
        } 
    }
}