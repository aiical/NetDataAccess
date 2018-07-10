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
    public class ThycGetDetail : CustomProgramBase
    {
        #region 入口函数 
        public bool Run(string parameters, IListSheet listSheet)
        {
            return GetShopDetail(parameters, listSheet);
        }
        #endregion

        #region GetShopDetail
        private bool GetShopDetail(string parameters, IListSheet listSheet)
        {
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            
            //输出excel表格包含的列
            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{ 
                "shopCode",
                "shopName",
                "cityCode",
                "cityName",
                "provinceCode",
                "provinceName",
                "level",
                "address",
                "lng",
                "lat",
                "serviceItems"});
             
            string exportDir = this.RunPage.GetExportDir();
             
            string resultFilePath = Path.Combine(exportDir, "途虎养车获取维修站详情.xlsx");

            //输出对象
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic);

            this.GetShopDetail(listSheet, pageSourceDir, resultEW);

            //保存到硬盘
            resultEW.SaveToDisk();

            return true;
        }
        #endregion 

        #region GetShopDetail
        /// <summary>
        /// GetShopDetail
        /// </summary>
        /// <param name="listSheet"></param>
        /// <param name="pageSourceDir"></param>
        /// <param name="resultEW"></param>
        private void GetShopDetail(IListSheet listSheet, string pageSourceDir, ExcelWriter resultEW)
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
                string shopCode = row["shopCode"];
                string shopName = row["shopName"];
                string level = "";
                string address = "";
                Nullable<decimal> lng = null;
                Nullable<decimal> lat = null;
                string serviceItems = "";

                string localFilePath = this.RunPage.GetFilePath(pageUrl, pageSourceDir);
                HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);

                HtmlNode levelNode = htmlDoc.DocumentNode.SelectSingleNode("//div[@class=\"shop-level\"]/span[1]");
                if (levelNode != null)
                {
                    level = levelNode.InnerText;
                }

                HtmlNode addressNode = htmlDoc.DocumentNode.SelectSingleNode("//div[@class=\"address clearfix\"]/div[@id=\"submitbtns\"]/span");
                if (addressNode != null)
                {
                    address = addressNode.InnerText;
                }

                HtmlNode scriptNode = htmlDoc.DocumentNode.SelectSingleNode("//div[@id=\"scriptSection\"]");
                if (scriptNode != null)
                {
                    string script = scriptNode.InnerText;
                    int lngBeginIndex = script.IndexOf("Position: '") + 11;
                    int lngEndIndex = script.IndexOf(",", lngBeginIndex);
                    int latBeginIndex = lngEndIndex + 1;
                    int latEndIndex = script.IndexOf("',", latBeginIndex);
                    lng = decimal.Parse(script.Substring(lngBeginIndex, lngEndIndex - lngBeginIndex));
                    lat = decimal.Parse(script.Substring(latBeginIndex, latEndIndex - latBeginIndex));
                }

                StringBuilder serviceItemSB = new StringBuilder();
                HtmlNodeCollection allServiceItemNodes = htmlDoc.DocumentNode.SelectNodes("//div[@class=\"sever-xm\"]/ul/li");

                if (allServiceItemNodes != null)
                {
                    for (int j = 0; j < allServiceItemNodes.Count; j++)
                    {
                        HtmlNode serviceNode = allServiceItemNodes[j];
                        if (!serviceNode.Attributes.Contains("class") || serviceNode.Attributes["class"].Value != "not-have")
                        {
                            string serviceText = serviceNode.InnerText.Trim();
                            serviceItemSB.Append(serviceText + ";");
                        }
                    }
                    serviceItems = serviceItemSB.ToString();
                }

                Dictionary<string, object> f2vs = new Dictionary<string, object>();
                f2vs.Add("provinceCode", provinceCode);
                f2vs.Add("provinceName", provinceName);
                f2vs.Add("cityCode", cityCode);
                f2vs.Add("cityName", cityName);
                f2vs.Add("shopCode", shopCode);
                f2vs.Add("shopName", shopName);
                f2vs.Add("level", level);
                f2vs.Add("address", address);
                f2vs.Add("lng", lng);
                f2vs.Add("lat", lat);
                f2vs.Add("serviceItems", serviceItems);
                resultEW.AddRow(f2vs);
            }
        }
        #endregion
    }
}