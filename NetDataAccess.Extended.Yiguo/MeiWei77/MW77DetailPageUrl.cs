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
    /// 美味七七
    /// 从本地html中获取并记录下商品详情信息
    /// </summary>
    public class MW77DetailPageUrl : CustomProgramBase
    {
        #region 入口函数
        public bool Run(string parameters, IListSheet listSheet)
        {
            return GetAllDetailPageUrl(listSheet);
        }
        #endregion

        #region 获取并记录下商品详情信息
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
            resultColumnDic.Add("productCode", 5);
            resultColumnDic.Add("productName", 6);
            resultColumnDic.Add("productCurrentPrice", 7);
            resultColumnDic.Add("productOldPrice", 8); 
            resultColumnDic.Add("category1Code", 9);
            resultColumnDic.Add("category2Code", 10);
            resultColumnDic.Add("category3Code", 11);
            resultColumnDic.Add("category1Name", 12);
            resultColumnDic.Add("category2Name", 13);
            resultColumnDic.Add("category3Name", 14);
            resultColumnDic.Add("standard", 15);
            string resultFilePath = Path.Combine(exportDir, this.RunPage.Project.Name + "_AllDetailPageUrl.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic);

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
                    string category1Code = row["category1Code"];
                    string category2Code = row["category2Code"];
                    string category3Code = row["category3Code"];
                    string category1Name = row["category1Name"];
                    string category2Name = row["category2Name"];
                    string category3Name = row["category3Name"];
                    string detailPageUrlPrefix = "http://www.yummy77.com"; 
                    string localFilePath = this.RunPage.GetFilePath(url, pageSourceDir);

                    try
                    {
                        HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                        HtmlNode listNode = htmlDoc.DocumentNode.SelectSingleNode("//div[@class=\"productlist\"]");
                        if (listNode != null)
                        {
                            //HtmlNodeCollection allPageNodes = listNode.SelectNodes("./div[@class='p_item_container p_item_ab ']");
                            HtmlNodeCollection allPageNodes = listNode.SelectNodes("./div");
                            if (allPageNodes != null)
                            {
                                foreach (HtmlNode pagesNode in allPageNodes)
                                {
                                    HtmlNodeCollection pageItemList = pagesNode.SelectNodes("./div");
                                    foreach (HtmlNode pageNode in pageItemList)
                                    {
                                        string productCode = "";
                                        string productName = "";
                                        string productCurrentPrice = "";
                                        string productOldPrice = "";
                                        string detailPageUrl = "";
                                        string detailPageName = "";
                                        string standard = "";

                                        HtmlNode nameNode = pageNode.SelectSingleNode("./span[@class=\"pname_div\"]/a");
                                        detailPageUrl = detailPageUrlPrefix + nameNode.Attributes["href"].Value;
                                        int startIndex = detailPageUrl.LastIndexOf("/") + 1;
                                        int endIndex = detailPageUrl.LastIndexOf(".");
                                        int length = endIndex - startIndex;

                                        //商品类型为礼品卡时，length==0，不用获取详情页
                                        if (length > 0)
                                        {
                                            detailPageName = detailPageUrl.Substring(startIndex, length);
                                            productCode = detailPageName;
                                            productName = nameNode.InnerText.Trim();

                                            HtmlNode pcNode = pageNode.SelectSingleNode("./span[@class=\"price_div\"]/span[@class=\"pcprice_sp\"]");
                                            if (pcNode != null)
                                            {
                                                productCurrentPrice = pcNode.InnerText.Trim().Substring(1);
                                            }

                                            HtmlNode pmNode = pageNode.SelectSingleNode("./span[@class=\"price_div\"]/span[@class=\"pmprice_sp\"]");
                                            if (pmNode != null)
                                            {
                                                productOldPrice = pmNode.InnerText.Trim().Substring(1);
                                            }

                                            HtmlNode standardNode = pageNode.SelectSingleNode("./div[@class=\"p_item_mark\"]/ul/li[@_pid=\"" + productCode + "\"]");
                                            if (standardNode != null)
                                            {
                                                standard = standardNode.InnerText.Trim();
                                            }

                                            if (!allProductCodes.ContainsKey(productCode))
                                            {
                                                allProductCodes.Add(productCode, null);
                                                Dictionary<string, string> p2vs = new Dictionary<string, string>();
                                                p2vs.Add("detailPageUrl", detailPageUrl);
                                                p2vs.Add("detailPageName", detailPageName);
                                                p2vs.Add("productCode", productCode);
                                                p2vs.Add("productName", productName);
                                                p2vs.Add("productCurrentPrice", productCurrentPrice);
                                                p2vs.Add("productOldPrice", productOldPrice);
                                                p2vs.Add("category1Code", category1Code);
                                                p2vs.Add("category2Code", category2Code);
                                                p2vs.Add("category3Code", category3Code);
                                                p2vs.Add("category1Name", category1Name);
                                                p2vs.Add("category2Name", category2Name);
                                                p2vs.Add("category3Name", category3Name);
                                                p2vs.Add("standard", standard);
                                                resultEW.AddRow(p2vs);
                                            }
                                        }
                                    }
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
            TaskManager.StartTask("易果", "美味77获取所有详情页", resultFilePath, null, null, false);
            
            return true;
        }
        #endregion
    }
}