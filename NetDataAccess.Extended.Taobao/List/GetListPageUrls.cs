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

namespace NetDataAccess.Extended.Taobao.List
{
    /// <summary>
    /// GetListPageUrls
    /// </summary>
    [System.Runtime.InteropServices.ComVisibleAttribute(true)]
    public class GetListPageUrls : ExternalRunWebPage
    {
        private string _ShopSearchPageUrlFormat = "https://shopsearch.taobao.com/search?app=shopsearch&q={keyword}&s={pageIndex}";
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

        private int _PageCount = 100;
        private int PageCount
        {
            get
            {
                return this._PageCount;
            }
        }

        private string GetShopSearchPageUrl(string keyword, int pageNum)
        {
            int pageIndex = pageNum * 20;
            string pageUrl = this.ShopSearchPageUrlFormat.Replace("{keyword}", keyword).Replace("{pageIndex}", pageIndex.ToString());
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
                    "keyword",
                    "pageNum"});
            string slPath = Path.Combine(exportDir, "淘宝店铺列表页.xlsx");
            ExcelWriter slEW = new ExcelWriter(slPath, "List", columnDic, null);

            string[] allKeywords = this.GetAllKeywords();
            Dictionary<string, string> keywordDic = new Dictionary<string, string>();            
            for (int i = 0; i < allKeywords.Length; i++)
            {
                string keyword = allKeywords[i].Trim();
                if (!CommonUtil.IsNullOrBlank(keyword) && !keywordDic.ContainsKey(keyword))
                {
                    keywordDic.Add(keyword, null);
                    for (int pageNum = 0; pageNum < PageCount; pageNum++)
                    {
                        string pageUrl = this.GetShopSearchPageUrl(keyword, pageNum);
                        Dictionary<string, string> row = new Dictionary<string, string>();
                        row.Add("detailPageUrl", pageUrl);
                        row.Add("detailPageName", keyword + "_" + pageNum.ToString());
                        row.Add("keyword", keyword);
                        row.Add("pageNum", (pageNum + 1).ToString());
                        slEW.AddRow(row);
                    }
                }
            }
            slEW.SaveToDisk();
            return succeed;
        }
    }
}