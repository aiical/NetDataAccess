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

namespace NetDataAccess.Extended.Id.List
{
    /// <summary>
    /// GetListPageUrls
    /// </summary>
    [System.Runtime.InteropServices.ComVisibleAttribute(true)]
    public class GetDetailPageUrls : ExternalRunWebPage
    {
        private string _ShopSearchPageUrlFormat = "https://idempiere.atlassian.net/si/jira.issueviews:issue-xml/IDEMPIERE-{pageIndex}/IDEMPIERE-{pageIndex}.xml";
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
        private string GetShopSearchPageUrl(int pageIndex)
        { 
            string pageUrl = this.ShopSearchPageUrlFormat.Replace("{pageIndex}", pageIndex.ToString());
            return pageUrl;
        }

        public override bool AfterAllGrab(IListSheet listSheet)
        {
            return this.GenerateUrlListFile();
        }

        private int GetTotalPageCount()
        {
            string[] strs = this.Parameters.Split(new string[] { "," }, StringSplitOptions.None);
            return int.Parse(strs[0]);
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
                    "giveUpGrab" });
            string slPath = Path.Combine(exportDir, "Id详情页.xlsx");
            ExcelWriter slEW = new ExcelWriter(slPath, "List", columnDic, null);

            int totalPageCount = this.GetTotalPageCount();
            int pageIndex = 1;
            while (pageIndex <= totalPageCount)
            { 
                string pageUrl = this.GetShopSearchPageUrl(pageIndex);
                Dictionary<string, string> row = new Dictionary<string, string>();
                row.Add("detailPageUrl", pageUrl);
                row.Add("detailPageName", pageIndex.ToString());  
                slEW.AddRow(row);
                pageIndex = pageIndex + 1;

            }
            slEW.SaveToDisk();
            return succeed;
        }
    }
}