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
using NetDataAccess.Base.Server;

namespace NetDataAccess.Extended.Yiguo
{
    /// <summary>
    /// 本来生活
    /// 获取并输出本来生活所有商品详情页地址
    /// </summary>
    public class BenlaiDetailPageUrl : CustomProgramBase
    {
        #region 入口函数
        public bool Run(string parameters, IListSheet listSheet)
        {
            return GetAllDetailPageUrl(listSheet);
        }
        #endregion

        #region 获取并输出本来生活所有商品详情页地址
        private bool GetAllDetailPageUrl(IListSheet listSheet)
        {
            string exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("productSysNo", 5);
            resultColumnDic.Add("productName", 6);
            resultColumnDic.Add("productPromotionWord", 7);
            resultColumnDic.Add("productCurrentPrice", 8);
            resultColumnDic.Add("productOldPrice", 9);
            resultColumnDic.Add("canDelivery", 10); 
            resultColumnDic.Add("category1Code", 11);
            resultColumnDic.Add("category2Code", 12);
            resultColumnDic.Add("category3Code", 13);
            resultColumnDic.Add("category1Name", 14);
            resultColumnDic.Add("category2Name", 15);
            resultColumnDic.Add("category3Name", 16);
            resultColumnDic.Add("district", 17);
            string resultFilePath = Path.Combine(exportDir, "本来生活获取所有详情页.xlsx");
            
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic);

            Dictionary<string, string> goodsDic = new Dictionary<string, string>();

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;
            string categoryNameColumnName = SysConfig.DetailPageNameFieldName;

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                { 
                    string url = row[detailPageUrlColumnName]; 
                    string category1Code = row["category1Code"];
                    string category2Code = row["category2Code"];
                    string category3Code = row["category3Code"];
                    string category1Name = row["category1Name"];
                    string category2Name = row["category2Name"];
                    string category3Name = row["category3Name"];
                    string district = row["district"];
                    string cookie = row["cookie"];
                    string detailPageUrlPrefix = "http://www.benlai.com";
                    string prefix = "";
                    switch (district)
                    {
                        case "华东":
                            prefix = "/huadong/item";
                            break;
                        case "华北":
                            prefix = "/item";
                            break;
                        default:
                            break;
                    }
                    string localFilePath = this.RunPage.GetFilePath(url, pageSourceDir);
                    TextReader tr = null;

                    try
                    {
                        tr = new StreamReader(localFilePath);
                        //获取到的文件内容为json对象
                        /*
                         {
	                        "TotalProductNum":1,
	                        "TotalNum":1,
	                        "SelectNum":1,
	                        "ProductList":[
		                        {
			                        "ProductSysNo":45790,
			                        "ProductName":"飞利浦HD3031/05 飞利浦电饭煲",
			                        "ProductPromotionWord":"HD3031/05 米饭更美味",
			                        "ProductLink":"/item-45790.html",
			                        "ProductImageLink":"http://image1.benlailife.com/ProductImages/000/000/045/790/medium/9cf3b6e9-785a-487a-9ade-5187d2396d44.jpg",
			                        "ProductAreaGroup":26,
			                        "ProductCurrentPrice":599.000000,
			                        "ProductCDPrice":-1,
			                        "ProductNowPrice":599.000000,
			                        "ProductOldPrice":-999999,
			                        "ProductCDQty":-999999,
			                        "ProductSaleQty":-999999,
			                        "Desc":0,
			                        "IsCanDelivery":1,
			                        "Inventory":20,
			                        "ProductStatus":0,
			                        "Status":0,
			                        "Tag1":"新品",
			                        "Tag2":null,
			                        "Tag3":""
		                        }
	                        ]
                        }
                        */

                        string webPageJson = tr.ReadToEnd();

                        //解析json对象
                        JObject rootJo = JObject.Parse(webPageJson);
                        
                        jt = rootJo.SelectToken("ProductList");
                        JArray jas = jt as JArray;

                        foreach (JObject jo in jas)
                        {
                            if (jo["ProductLink"].ToString().StartsWith(prefix))
                            {
                                string detailPageUrl = detailPageUrlPrefix + jo["ProductLink"].ToString();
                                string productSysNo = jo["ProductSysNo"].ToString();
                                //string detailPageName = district + "_" + category1Name + "_" + category2Name + "_" + category3Name + "_" + productSysNo;
                                string detailPageName = district + "_" + category1Name + "_" + category2Name + "_" + productSysNo;
                                string productName = jo["ProductName"].ToString();
                                string productPromotionWord = jo["ProductPromotionWord"].ToString();
                                string productCurrentPrice = jo["ProductCurrentPrice"].ToString();
                                string productOldPrice = jo["ProductOldPrice"].ToString();
                                string canDelivery = jo["IsCanDelivery"].ToString();
                                if (!goodsDic.ContainsKey(detailPageName))
                                {
                                    goodsDic.Add(detailPageName, null);
                                    Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                    f2vs.Add("detailPageUrl", detailPageUrl);
                                    f2vs.Add("detailPageName", detailPageName);
                                    f2vs.Add("cookie", cookie);
                                    f2vs.Add("productName", productName);
                                    f2vs.Add("productPromotionWord", productPromotionWord);
                                    f2vs.Add("productCurrentPrice", productCurrentPrice);
                                    f2vs.Add("productOldPrice", productOldPrice == "-999999" ? "" : productOldPrice);
                                    f2vs.Add("canDelivery", canDelivery);
                                    f2vs.Add("category1Code", category1Code);
                                    f2vs.Add("category2Code", category2Code);
                                    f2vs.Add("category3Code", category3Code);
                                    f2vs.Add("category1Name", category1Name);
                                    f2vs.Add("category2Name", category2Name);
                                    f2vs.Add("category3Name", category3Name);
                                    f2vs.Add("district", district);
                                    f2vs.Add("productSysNo", productSysNo);
                                    resultEW.AddRow(f2vs);
                                }
                            }
                            else
                            {
                                throw new Exception("页面获取错误，站点不同. Url=" + url);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        if (tr != null)
                        {
                            tr.Dispose();
                            tr = null;
                        }
                        this.RunPage.InvokeAppendLogText("读取出错.  " + ex.Message + " LocalPath = " + localFilePath, LogLevelType.Error, true);
                        throw ex;
                    }
                }
            }
            resultEW.SaveToDisk();

            return true;
        }
        #endregion
    }
}