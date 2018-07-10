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

namespace NetDataAccess.Extended.Jingdong.List
{
    /// <summary>
    /// GetShopKeywordList
    /// </summary>
    [System.Runtime.InteropServices.ComVisibleAttribute(true)]
    public class GetShopKeywordList : ExternalRunWebPage
    {  
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            return this.GenerateShopKeywordListFile();
        }

        private string[] GetAllKeywords()
        {
            string[] allKeywords = this.Parameters.Split(new string[] { "," }, StringSplitOptions.None);
            return allKeywords;
        } 

        /// <summary>
        /// 生成店铺关键字列表文件
        /// </summary>
        /// <param name="listSheet"></param>
        /// <returns></returns>
        private bool GenerateShopKeywordListFile()
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
                    "keyword"});
            string slPath = Path.Combine(exportDir, "京东关键字店铺.xlsx");
            ExcelWriter slEW = new ExcelWriter(slPath, "List", columnDic, null);

            string[] allKeywords = this.GetAllKeywords();
            Dictionary<string, string> keywordDic = new Dictionary<string, string>();            
            for (int i = 0; i < allKeywords.Length; i++)
            {
                string keyword = allKeywords[i].Trim();
                if (!CommonUtil.IsNullOrBlank(keyword) && !keywordDic.ContainsKey(keyword))
                {
                    keywordDic.Add(keyword, null);
                    Dictionary<string, string> row = new Dictionary<string, string>();
                    row.Add("detailPageUrl", keyword);
                    row.Add("detailPageName", keyword);
                    row.Add("keyword", keyword); 
                    slEW.AddRow(row);
                }
            }
            slEW.SaveToDisk();
            return succeed;
        }
    }
}