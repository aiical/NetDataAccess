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
    /// 天天果园
    /// 获取所有商品详情页地址
    /// </summary>
    public class TTGYDetailPageUrl : CustomProgramBase
    {
        #region 入口函数
        public bool Run(string parameters, IListSheet listSheet)
        {
            bool succeed = GetAllDetailPageUrl(listSheet);
            return succeed;
        }
        #endregion

        #region 获取所有详情页地址
        private bool GetAllDetailPageUrl(IListSheet listSheet)
        {
            string exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            string[] resultColumns = new string[]{ "detailPageUrl",
                "detailPageName",
                "cookie", 
                "grabStatus", 
                "giveUpGrab",
                "productCode", 
                "productName",
                "productCurrentPrice",
                "productOldPrice",
                "categoryCode", 
                "categoryName", 
                "standard", 
                "city"};
            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(resultColumns);
            string resultFilePath = Path.Combine(exportDir, this.RunPage.Project.Name + "_AllDetailPageUrl.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic);

            string detailPageUrlPrefix = "http://www.fruitday.com";
            Dictionary<string, string> allProductCodes = new Dictionary<string, string>();

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;
            string categoryNameColumnName = SysConfig.DetailPageNameFieldName;

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                { 
                    string url = row[detailPageUrlColumnName];
                    string categoryCode = row["categoryCode"];
                    string categoryName = row["categoryName"];
                    string cookie = row["cookie"];
                    string city = row["city"];  
                    string localFilePath = this.RunPage.GetFilePath(url, pageSourceDir);

                    try
                    {
                        HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                        HtmlNodeCollection allItemNodes = htmlDoc.DocumentNode.SelectNodes("//div[@class=\"leftpart pull-left\"]/ul/li");
                        if (allItemNodes != null)
                        {
                            foreach (HtmlNode itemNode in allItemNodes)
                            {
                                string productCode = "";
                                string productName = "";
                                string productCurrentPrice = "";
                                string productOldPrice = "";
                                string detailPageUrl = "";
                                string detailPageName = "";
                                string standard = "";

                                HtmlNode urlNode = itemNode.SelectSingleNode("./div/div[@class=\"s-img\"]/a");
                                detailPageUrl = detailPageUrlPrefix + urlNode.Attributes["href"].Value;
                                int startIndex = detailPageUrl.LastIndexOf("/") + 1;
                                detailPageName = detailPageUrl.Substring(startIndex);
                                productCode = detailPageName;

                                HtmlNodeCollection propertyNodes = itemNode.SelectSingleNode("./div/div[@class=\"s-info clearfix\"]").ChildNodes;
                                foreach (HtmlNode propertyNode in propertyNodes)
                                {
                                    if (propertyNode.NodeType == HtmlNodeType.Text)
                                    {
                                        productName = propertyNode.InnerText.Trim();
                                    }
                                    else
                                    {
                                        if (propertyNode.Attributes.Contains("class")
                                            && propertyNode.Attributes["class"].Value == "s-unit pull-right font-color")
                                        {
                                            string priceStr = propertyNode.InnerText.Trim();
                                            productCurrentPrice = priceStr.Substring(1);
                                        }
                                    }
                                }

                                HtmlNode standardNode = itemNode.SelectSingleNode("./div/div[@class=\"p-operate clearfix\"]");
                                if (standardNode != null)
                                {
                                    standard = standardNode.InnerText.Trim();
                                }
                                detailPageName = city + "_" + detailPageName;

                                if (!allProductCodes.ContainsKey(detailPageName))
                                {
                                    allProductCodes.Add(detailPageName, null);
                                    Dictionary<string, string> p2vs = new Dictionary<string, string>();
                                    p2vs.Add("detailPageUrl", detailPageUrl + "?city=" + city);
                                    p2vs.Add("detailPageName", detailPageName);
                                    p2vs.Add("city", city);
                                    p2vs.Add("cookie", cookie);
                                    p2vs.Add("productCode", productCode);
                                    p2vs.Add("productName", productName);
                                    p2vs.Add("productCurrentPrice", productCurrentPrice);
                                    p2vs.Add("productOldPrice", productOldPrice);
                                    p2vs.Add("categoryCode", categoryCode);
                                    p2vs.Add("categoryName", categoryName);
                                    p2vs.Add("standard", standard);
                                    resultEW.AddRow(p2vs);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        this.RunPage.InvokeAppendLogText("读取出错.  " + ex.Message + " LocalPath = " + localFilePath, LogLevelType.Error, true);
                        throw ex;
                    }
                }
            }
            resultEW.SaveToDisk();

            //执行后续任务
            TaskManager.StartTask("易果", "天天果园获取所有详情页", resultFilePath, null, null, false);
            return true;
        }
        #endregion
    }
}