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

namespace NetDataAccess.Extended.Lvsejianzhu
{
    /// <summary>
    /// GetListPageUrls
    /// </summary>
    [System.Runtime.InteropServices.ComVisibleAttribute(true)]
    public class GetListPageUrls : ExternalRunWebPage
    {
        private string _ShopSearchPageUrlFormat = "http://www.gbmap.org/main2.php?diquselectinfo=,1,&zdanweiselectinfo=&sdanweiselectinfo=&tdanweiselectinfo=&leibieselectinfo=&gongnengselectinfo=&jishuselectinfo=&shijianselectinfo=&projectid=&page={pageIndex}";
        private string ShopSearchPageUrlFormat
        {
            get
            {
                return _ShopSearchPageUrlFormat;
            }
            set
            {
                _ShopSearchPageUrlFormat = value;
            }
        }

        private int _PageCount = 170;
        private int PageCount
        {
            get
            {
                return this._PageCount;
            }
        }

        private string GetShopSearchPageUrl(int pageNum)
        {
            int pageIndex = pageNum + 1;
            string pageUrl = this.ShopSearchPageUrlFormat.Replace("{pageIndex}", pageIndex.ToString());
            return pageUrl;
        }

        public override bool AfterAllGrab(IListSheet listSheet)
        {
            return this.GenerateUrlListFile();
        }

        private string[] GetAllKeywords()
        {
            string[] allKeywords = this.Parameters.Split(new string[] { "," }, StringSplitOptions.None);
            return allKeywords;
        } 

        /// <summary>
        /// 生成车辆详细信息
        /// </summary>
        /// <param name="listSheet"></param>
        /// <returns></returns>
        private bool GenerateUrlListFile()
        {
            bool succeed = true;
            string exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> columnDic = CommonUtil.InitStringIndexDic(new string[]{
                    "detailPageUrl",
                    "detailPageName", 
                    "cookie",
                    "grabStatus", 
                    "giveUpGrab",
                    "pageNum"});
            string slPath = Path.Combine(exportDir, "绿色建筑列表页.xlsx");
            ExcelWriter slEW = new ExcelWriter(slPath, "List", columnDic, null);

            for (int pageNum = 0; pageNum < PageCount; pageNum++)
            {
                string pageUrl = this.GetShopSearchPageUrl(pageNum);
                Dictionary<string, string> row = new Dictionary<string, string>();
                row.Add("detailPageUrl", pageUrl);
                row.Add("detailPageName", pageNum.ToString());
                row.Add("pageNum", (pageNum + 1).ToString());
                slEW.AddRow(row);
            }
            slEW.SaveToDisk();
            return succeed;
        }
    }
}