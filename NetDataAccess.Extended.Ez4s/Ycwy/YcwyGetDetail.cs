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
    public class YcwyGetDetail : CustomProgramBase
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
                "provinceName",
                "serviceTime",
                "tel",
                "address",
                "lng",
                "lat",
                "serviceItems"});

            string exportDir = this.RunPage.GetExportDir();

            Dictionary<string, string> resultColumnFormat = new Dictionary<string, string>(); 
            resultColumnFormat.Add("lat", "#,##0.000000");
            resultColumnFormat.Add("lng", "#,##0.000000"); 
             
            string resultFilePath = Path.Combine(exportDir, "养车无忧维修站详情.xlsx");

            //输出对象
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, resultColumnFormat);

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
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string pageUrl = row[SysConfig.DetailPageUrlFieldName];
                    string provinceName = row["provinceName"];
                    string cityCode = row["cityCode"];
                    string cityName = row["cityName"];
                    string shopCode = row["shopCode"];
                    string shopName = row["shopName"];
                    string serviceTime = "";
                    string tel = "";
                    string address = "";
                    Nullable<decimal> lng = null;
                    Nullable<decimal> lat = null;
                    string serviceItems = "";

                    string localFilePath = this.RunPage.GetFilePath(pageUrl, pageSourceDir);
                    HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);

                    HtmlNodeCollection allInfoNameNodes = htmlDoc.DocumentNode.SelectNodes("//div[@class=\"shopInfo\"]/dl/dt");
                    if (allInfoNameNodes != null)
                    {
                        foreach (HtmlNode infoNameNode in allInfoNameNodes)
                        {
                            string infoName = infoNameNode.InnerText;
                            if (infoName.StartsWith("服务时间"))
                            {
                                HtmlNode infoNode = HtmlDocumentHelper.GetNextNode(infoNameNode, "dd");
                                if (infoNode != null)
                                {
                                    serviceTime = infoNode.InnerText;
                                }
                            }
                            /*
                            else if (infoName.StartsWith("服务地址"))
                            {
                                HtmlNode infoNode = HtmlDocumentHelper.GetNextNode(infoNameNode, "dd");
                                if (infoNode != null)
                                {
                                    HtmlNode infoSpanNode = infoNode.SelectSingleNode("./span");
                                    if (infoSpanNode != null)
                                    {
                                        address = infoSpanNode.InnerText;
                                    }
                                }
                            } 
                            else if (infoName.StartsWith("服务电话"))
                            {
                                HtmlNode infoNode = HtmlDocumentHelper.GetNextNode(infoNameNode, "dd");
                                if (infoNode != null)
                                {
                                    HtmlNode infoSpanNode = infoNode.SelectSingleNode("./span");
                                    if (infoSpanNode != null)
                                    {
                                        tel = infoSpanNode.InnerText;
                                    }
                                }
                            }
                            */
                        }
                    }

                    HtmlNodeCollection allScriptNodes = htmlDoc.DocumentNode.SelectNodes("//script");
                    if (allScriptNodes != null)
                    {
                        foreach (HtmlNode scriptNode in allScriptNodes)
                        {
                            string script = scriptNode.InnerText;
                            if (script.Contains("var lng = \""))
                            {
                                int lngBeginIndex = script.IndexOf("var lng = \"") + 11;
                                int lngEndIndex = script.IndexOf("\";", lngBeginIndex);
                                int latBeginIndex = script.IndexOf("var lat = \"") + 11;
                                int latEndIndex = script.IndexOf("\";", latBeginIndex);
                                lng = decimal.Parse(script.Substring(lngBeginIndex, lngEndIndex - lngBeginIndex));
                                lat = decimal.Parse(script.Substring(latBeginIndex, latEndIndex - latBeginIndex));
                                break;
                            }
                        }
                    }

                    HtmlNode telNode = htmlDoc.DocumentNode.SelectSingleNode("//input[@id=\"HiddenStrPhone\"]");
                    if (telNode != null)
                    {
                        tel = telNode.Attributes["value"].Value;
                    }

                    HtmlNode keyWordNode = htmlDoc.DocumentNode.SelectSingleNode("//meta[@name='keywords']");
                    string keyWord = keyWordNode.Attributes["content"].Value;
                    string[] kws = keyWord.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                    address = kws[1];

                    StringBuilder serviceItemSB = new StringBuilder();
                    HtmlNodeCollection allServiceGroupNodes = htmlDoc.DocumentNode.SelectNodes("//div[@class=\"shopProject\"]/table/tbody/tr");

                    if (allServiceGroupNodes != null)
                    {
                        for (int j = 0; j < allServiceGroupNodes.Count; j++)
                        {
                            HtmlNode serviceGroupNode = allServiceGroupNodes[j];
                            HtmlNodeCollection allServiceNodes = serviceGroupNode.SelectNodes("./td");
                            foreach (HtmlNode serviceNode in allServiceNodes)
                            {
                                string serviceText = serviceNode.InnerText.Trim();
                                if (serviceText.EndsWith("："))
                                {
                                    serviceItemSB.Append(serviceText);
                                }
                                else
                                {
                                    serviceItemSB.Append(serviceText + ";");
                                }
                            }
                        }
                        serviceItems = serviceItemSB.ToString();
                    }

                    Dictionary<string, object> f2vs = new Dictionary<string, object>();
                    f2vs.Add("provinceName", provinceName);
                    f2vs.Add("cityCode", cityCode);
                    f2vs.Add("cityName", cityName);
                    f2vs.Add("shopCode", shopCode);
                    f2vs.Add("shopName", shopName);
                    f2vs.Add("serviceTime", serviceTime);
                    f2vs.Add("address", address);
                    f2vs.Add("lng", lng);
                    f2vs.Add("lat", lat);
                    f2vs.Add("tel", tel);
                    f2vs.Add("serviceItems", serviceItems);
                    resultEW.AddRow(f2vs);
                }
            }
        }
        #endregion
    }
}