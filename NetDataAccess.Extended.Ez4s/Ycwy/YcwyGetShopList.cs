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

namespace NetDataAccess.Extended.Ez4s
{
    public class YcwyGetShopList : CustomProgramBase
    {
        #region 入口函数 
        public bool Run(string parameters, IListSheet listSheet)
        {
            return GetShopList(parameters, listSheet);
        }
        #endregion

        #region GetShopList
        private bool GetShopList(string parameters, IListSheet listSheet)
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
                "provinceName", 
                "cityName",
                "cityCode",
                "shopCode",
                "shopName"});
             
            string exportDir = this.RunPage.GetExportDir();
             
            string resultFilePath = Path.Combine(exportDir, "养车无忧详情页.xlsx");

            //输出对象
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic);

            //从下载到的首页html中，获取列表页个数，并形成所有列表页url
            GetShopList(listSheet, pageSourceDir, resultEW);

            //保存到硬盘
            resultEW.SaveToDisk();

            return true;
        }
        #endregion

        #region GetShopList
        /// <summary>
        /// GetCities
        /// </summary>
        /// <param name="listSheet"></param>
        /// <param name="pageSourceDir"></param>
        /// <param name="resultEW"></param>
        private void GetShopList(IListSheet listSheet, string pageSourceDir, ExcelWriter resultEW)
        {
            for (int i = 0; i < listSheet.RowCount; i++)
            { 
                string pageUrl = listSheet.PageUrlList[i];
                Dictionary<string, string> row = listSheet.GetRow(i);
                string provinceName = row["provinceName"];
                string cityName = row["cityName"];
                string cityCode = row["cityCode"];
                string localFilePath = this.RunPage.GetFilePath(pageUrl, pageSourceDir);
                HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);

                HtmlNodeCollection allShopNodes = htmlDoc.DocumentNode.SelectNodes("//div[@id=\"cityMapLeft\"]/div/b/a");

                for (int j = 0; j < allShopNodes.Count; j++)
                {
                    HtmlNode shopNode = allShopNodes[j];
                    string shopUrl = shopNode.Attributes["href"].Value;
                    string[] shopPieces = shopUrl.Split(new string[] { "/" }, StringSplitOptions.RemoveEmptyEntries);
                    string shopCodeStr = shopPieces[shopPieces.Length - 1];
                    string shopCode = shopCodeStr.Substring(0, shopCodeStr.IndexOf("."));
                    string shopName = shopNode.InnerText.Trim();

                    Dictionary<string, string> f2vs = new Dictionary<string, string>();
                    f2vs.Add("detailPageUrl", shopUrl);
                    f2vs.Add("detailPageName", shopCode + shopName);
                    f2vs.Add("provinceName", provinceName);
                    f2vs.Add("cityCode", cityCode);
                    f2vs.Add("cityName", cityName);
                    f2vs.Add("shopCode", shopCode);
                    f2vs.Add("shopName", shopName);
                    resultEW.AddRow(f2vs);
                }
            }
        }
        #endregion
    }
}