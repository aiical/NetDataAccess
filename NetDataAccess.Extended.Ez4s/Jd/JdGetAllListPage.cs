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

namespace NetDataAccess.Extended.Ez4s
{
    public class JdGetAllListPage : CustomProgramBase
    {
        #region 入口函数 
        public bool Run(string parameters, IListSheet listSheet)
        {
            return GetAllListPage(parameters, listSheet);
        }
        #endregion

        #region GetAllListPage
        private bool GetAllListPage(string parameters, IListSheet listSheet)
        {
            string serviceCatFilePath = parameters;

            //已经下载下来的首页html保存到的目录（文件夹）
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            
            //输出excel表格包含的列
            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{ 
                "cat1Name", 
                "cat2Code", 
                "cat2Name",
                "areaLevel1Code", 
                "areaLevel1Name",
                "areaLevel2Code", 
                "areaLevel2Name",
                "areaLevel3Code", 
                "areaLevel3Name",
                "shopCode",
                "shopName",
                "lng",
                "lat",
                "address",
                "skuName",
                "skuCode",
                "serviceTime",
                "tel",
                "price"});
             
            string exportDir = this.RunPage.GetExportDir();

            Dictionary<string, string> resultColumnFormat = new Dictionary<string, string>();
            resultColumnFormat.Add("price", "#,##0.00");
            resultColumnFormat.Add("lat", "#,##0.000000");
            resultColumnFormat.Add("lng", "#,##0.000000"); 

            string resultFilePath = Path.Combine(exportDir, "获取店铺服务所有列表页.xlsx");

            //输出对象
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, resultColumnFormat);

            //从下载到的首页html中，获取列表页个数，并形成所有列表页url
            GetAllListPage(listSheet, pageSourceDir, serviceCatFilePath, resultEW);

            //保存到硬盘
            resultEW.SaveToDisk();

            return true;
        }
        #endregion 

        #region GetAllListPage
        /// <summary>
        /// GetAllListPage
        /// </summary>
        /// <param name="listSheet"></param>
        /// <param name="pageSourceDir"></param>
        /// <param name="resultEW"></param>
        private void GetAllListPage(IListSheet listSheet, string pageSourceDir, string serviceCatFilePath, ExcelWriter resultEW)
        {
            Dictionary<string, string> shopServiceItemDic = new Dictionary<string, string>();
            for (int i = 0; i < listSheet.RowCount; i++)
            { 
                string pageUrl = listSheet.PageUrlList[i];
                Dictionary<string, string> row = listSheet.GetRow(i);
                string areaLevel1Code = row["areaLevel1Code"];
                string areaLevel1Name = row["areaLevel1Name"];
                string areaLevel2Code = row["areaLevel2Code"];
                string areaLevel2Name = row["areaLevel2Name"];
                string areaLevel3Code = row["areaLevel3Code"];
                string areaLevel3Name = row["areaLevel3Name"];
                string cat1Name = row["cat1Name"];
                string cat2Name = row["cat2Name"];
                string cat2Code = row["cat2Code"];

                string localFilePath = this.RunPage.GetFilePath(pageUrl, pageSourceDir);
                string fileText = FileHelper.GetTextFromFile(localFilePath);

                int jsonBeginIndex = fileText.IndexOf("{");
                int jsonEndIndex = fileText.LastIndexOf("}");

                string jsonStr = fileText.Substring(jsonBeginIndex, jsonEndIndex - jsonBeginIndex + 1);
                JObject rootJo = JObject.Parse(jsonStr);
                JArray shopListArray= rootJo.SelectToken("data").SelectToken("shopList") as JArray;

                for (int j = 0; j < shopListArray.Count; j++)
                {
                    JObject shopObject = shopListArray[j] as JObject;

                    string shopCode = (shopObject.GetValue("jd_shopid") as JValue).Value.ToString();
                    string shopName = (shopObject.GetValue("name") as JValue).Value.ToString();
                    string latLngStr = (shopObject.GetValue("maps") as JValue).Value.ToString();
                    string[] latLngs = latLngStr.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                    Nullable<decimal> lng = decimal.Parse(latLngs[1]);
                    Nullable<decimal> lat = decimal.Parse(latLngs[0]);
                    string address = (shopObject.GetValue("address") as JValue).Value.ToString();
                    string skuName = (shopObject.GetValue("skuName") as JValue).Value.ToString();
                    string skuCode = (shopObject.GetValue("skuId") as JValue).Value.ToString();
                    string serviceTime = (shopObject.GetValue("businessHours") as JValue).Value.ToString();
                    JValue telValue = shopObject.GetValue("telephone") as JValue;
                    string tel = telValue == null || telValue.Value == null ? "" : telValue.Value.ToString();
                    Nullable<decimal> price = decimal.Parse((shopObject.GetValue("price") as JValue).Value.ToString());

                    string shopServiceId = shopCode + "_" + skuCode;

                    if (!shopServiceItemDic.ContainsKey(shopServiceId))
                    {
                        shopServiceItemDic.Add(shopServiceId, "");
                        Dictionary<string, object> f2vs = new Dictionary<string, object>();
                        f2vs.Add("areaLevel1Code", areaLevel1Code);
                        f2vs.Add("areaLevel1Name", areaLevel1Name);
                        f2vs.Add("areaLevel2Code", areaLevel2Code);
                        f2vs.Add("areaLevel2Name", areaLevel2Name);
                        f2vs.Add("areaLevel3Code", areaLevel3Code);
                        f2vs.Add("areaLevel3Name", areaLevel3Name);
                        f2vs.Add("cat1Name", cat1Name);
                        f2vs.Add("cat2Code", cat2Code);
                        f2vs.Add("cat2Name", cat2Name);
                        f2vs.Add("shopCode", shopCode);
                        f2vs.Add("shopName", shopName);
                        f2vs.Add("lng", lng);
                        f2vs.Add("lat", lat);
                        f2vs.Add("address", address);
                        f2vs.Add("skuName", skuName);
                        f2vs.Add("skuCode", skuCode);
                        f2vs.Add("serviceTime", serviceTime);
                        f2vs.Add("tel", tel);
                        f2vs.Add("price", price);
                        resultEW.AddRow(f2vs);
                    }
                }
            }
        }
        #endregion
    }
}