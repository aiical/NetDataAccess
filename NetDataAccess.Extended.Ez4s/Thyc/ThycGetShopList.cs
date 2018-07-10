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
    public class ThycGetShopList : CustomProgramBase
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
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            
            //输出excel表格包含的列
            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                "detailPageUrl",
                "detailPageName", 
                "cookie",
                "grabStatus", 
                "giveUpGrab",
                "provinceName",
                "provinceCode",
                "cityName",
                "cityCode",
                "shopCode",
                "shopName"});
             
            string exportDir = this.RunPage.GetExportDir();
             
            string resultFilePath = Path.Combine(exportDir, "途虎养车获取维修站详情.xlsx");

            //输出对象
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic);

            this.GetShopList(listSheet, pageSourceDir, resultEW);

            //保存到硬盘
            resultEW.SaveToDisk();

            return true;
        }
        #endregion 

        #region GetShopList
        /// <summary>
        /// GetShopList
        /// </summary>
        /// <param name="listSheet"></param>
        /// <param name="pageSourceDir"></param>
        /// <param name="resultEW"></param>
        private void GetShopList(IListSheet listSheet, string pageSourceDir, ExcelWriter resultEW)
        {
            Dictionary<string, string> shopDic = new Dictionary<string, string>();
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                //listSheet中只有一条记录
                Dictionary<string, string> row = listSheet.GetRow(i);
                string pageUrl = row[SysConfig.DetailPageUrlFieldName];
                string provinceCode = row["provinceCode"];
                string provinceName = row["provinceName"];
                string cityCode = row["cityCode"];
                string cityName = row["cityName"];
                string localFilePath = this.RunPage.GetFilePath(pageUrl, pageSourceDir); 
                HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);

                HtmlNodeCollection allShopNodes = htmlDoc.DocumentNode.SelectNodes("//div[@class=\"non-list\"]/div/div[@class=\"shop-name clearfix\"]/a[@class=\"carparname\"]");

                if (allShopNodes != null)
                {
                    for (int j = 0; j < allShopNodes.Count; j++)
                    {
                        HtmlNode shopNode = allShopNodes[j];
                        string shopUrl = shopNode.Attributes["href"].Value;
                        string[] shopUrlPieces = shopUrl.Split(new string[] { "/" }, StringSplitOptions.RemoveEmptyEntries);
                        string[] shopPageNamePieces = shopUrlPieces[shopUrlPieces.Length - 1].Split(new string[] { "." }, StringSplitOptions.RemoveEmptyEntries);
                        string shopCode = shopPageNamePieces[0];
                        string shopName = shopNode.Attributes["title"].Value;

                        if (!shopDic.ContainsKey(shopCode))
                        {
                            shopDic.Add(shopCode, "");
                            Dictionary<string, string> f2vs = new Dictionary<string, string>();
                            f2vs.Add("detailPageUrl", shopUrl);
                            f2vs.Add("detailPageName", shopCode + shopName);
                            f2vs.Add("provinceCode", provinceCode);
                            f2vs.Add("provinceName", provinceName);
                            f2vs.Add("cityCode", cityCode);
                            f2vs.Add("cityName", cityName);
                            f2vs.Add("shopCode", shopCode);
                            f2vs.Add("shopName", shopName);
                            resultEW.AddRow(f2vs);
                        }
                    }
                }
            }
        }
        #endregion
    }
}