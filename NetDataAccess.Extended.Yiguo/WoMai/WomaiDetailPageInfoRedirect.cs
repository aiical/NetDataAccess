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
using NetDataAccess.Base.Writer;
using NetDataAccess.Base.Reader;
using NetDataAccess.Base.CsvHelper;
using NetDataAccess.Base.DB;
using HtmlAgilityPack;

namespace NetDataAccess.Extended.Yiguo
{
    public class WomaiDetailPageInfoRedirect : CustomProgramBase
    {
        public bool Run(string parameters, IListSheet listSheet)
        {
            return this.GeneratePageInfo(listSheet);
        }

        private bool GeneratePageInfo(IListSheet listSheet)
        {
            bool succeed = true;
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            string exportDir = this.RunPage.GetExportDir();
            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("productSysNo", 5);
            resultColumnDic.Add("category1Code", 6);
            resultColumnDic.Add("category2Code", 7);
            resultColumnDic.Add("category1Name", 8);
            resultColumnDic.Add("category2Name", 9);
            resultColumnDic.Add("district", 10);
            resultColumnDic.Add("商品名称", 11);
            resultColumnDic.Add("品牌", 12);
            resultColumnDic.Add("净含量", 13);
            resultColumnDic.Add("产品毛重", 14);
            resultColumnDic.Add("产地", 15);
            resultColumnDic.Add("保质期", 16);
            resultColumnDic.Add("规格", 17);
            resultColumnDic.Add("营销栏目", 18);
            resultColumnDic.Add("商品页URL", 19);
            string resultFilePath = Path.Combine(exportDir, "我买网获取所有详情页价格.xlsx");

            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic);

            GetPricePageUrl(listSheet, pageSourceDir, resultEW);

            resultEW.SaveToDisk();

            return succeed;
        }

        private void GetPricePageUrl(IListSheet listSheet, string pageSourceDir, ExcelWriter resultEW)
        {
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string pageUrl = listSheet.PageUrlList[i];
                string pageName = listSheet.PageNameList[i];
                string localFilePath = this.RunPage.GetFilePath(pageUrl, pageSourceDir);
                string productSysNo = row["productSysNo"];
                string category1Code = row["category1Code"];
                string category2Code = row["category2Code"];
                string category1Name = row["category1Name"];
                string category2Name = row["category2Name"];
                string district = row["district"];
                string cookie = row["cookie"];
                string name = ""; 
                string pinpai = "";
                string jinghanliang = "";
                string chanpinmaozhong = "";
                string chandi = "";
                string yingxiaolanmu = "";
                string baozhiqi = "";
                string guige = "";

                TextReader tr = null;

                try
                {
                    tr = new StreamReader(localFilePath, Encoding.GetEncoding("GBK"));
                    string webPageHtml = tr.ReadToEnd().Trim();
                    
                    HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                    htmlDoc.LoadHtml(webPageHtml);

                    if (pageUrl.Contains("http://kj."))
                    {
                        yingxiaolanmu = "我买海淘";
                    }
                    else if (pageUrl.Contains("http://zs."))
                    {
                        yingxiaolanmu = "包邮直送";
                    }
                    else if (pageUrl.Contains("http://jiu."))
                    {
                        yingxiaolanmu = "我卖酒";
                    }
                    else if (pageUrl.Contains("http://tuan.") || pageUrl.Contains("/tuan/"))
                    {
                        yingxiaolanmu = "团购";
                    }
                    else if (pageUrl.Contains("/shan/"))
                    {
                        yingxiaolanmu = "闪购";
                    }

                    if (pageUrl.Contains("http://tuan.") || pageUrl.Contains("/tuan/"))
                    {
                        //团购
                        if (!webPageHtml.Contains("您所查找的团购活动未上线"))
                        {
                            HtmlNode nameNode = htmlDoc.DocumentNode.SelectSingleNode("//a[@class=\"current\"]");
                            name = nameNode.InnerText.Trim();
                        }
                        else
                        {
                            continue;
                        }
                    }
                    else if (pageUrl.Contains("/shan/"))
                    {
                        //闪购 
                        HtmlNode nameNode = htmlDoc.DocumentNode.SelectSingleNode("//div[@class=\"detail-top\"]/div[1]/div[1]/img[1]");
                        name = nameNode.Attributes["alt"].Value;
                    }
                    else if (pageUrl.Contains("grandcru"))
                    {
                        //红酒
                        HtmlNode nameNode = htmlDoc.DocumentNode.SelectSingleNode("//label[@class=\"WrapTit\"]");
                        name = nameNode.InnerText.Trim();
                    }
                    else
                    {
                        HtmlNode nameNode = htmlDoc.DocumentNode.SelectSingleNode("//a[@class=\"last\"]");
                        name = nameNode.InnerText.Trim();

                        //品牌、净含量、产品毛重、产地、保质期、规格
                        HtmlNodeCollection allPropertyNodes = htmlDoc.DocumentNode.SelectNodes("//div[@class=\"detail_tab_pro_info\"]/ul/li");
                        if (allPropertyNodes != null)
                        {
                            foreach (HtmlNode propertyNode in allPropertyNodes)
                            {
                                string propertyText = propertyNode.InnerText;
                                HtmlAttribute titleAttr = propertyNode.Attributes["title"];
                                if (titleAttr != null)
                                {
                                    string propertyValue = titleAttr.Value;
                                    if (propertyText.Contains("品牌"))
                                    {
                                        pinpai = propertyValue;
                                    }
                                    else if (propertyText.Contains("净含量"))
                                    {
                                        jinghanliang = propertyValue;
                                    }
                                    else if (propertyText.Contains("产品毛重") || propertyText.Contains("商品毛重"))
                                    {
                                        chanpinmaozhong = propertyValue;
                                    }
                                    else if (propertyText.Contains("产地"))
                                    {
                                        chandi = propertyValue;
                                    }
                                    else if (propertyText.Contains("保质期"))
                                    {
                                        baozhiqi = propertyValue;
                                    }
                                    else if (propertyText.Contains("规格"))
                                    {
                                        guige = propertyValue;
                                    }
                                }
                            }
                        }
                    }

                    string priceUrl = "http://price.womai.com/PriceServer/open/productlist.do?ids=" + productSysNo + "&mid=0&usergroupid=100&properties=productLabels&prices=buyPrice%2CmarketPrice%2CWMPrice%2CVIPPrice%2CsaveMoney%2CuserPoints%2CspecialPrice&defData=n&t=0.041262077167630196&callback=jsonp1457471312047";

                    Dictionary<string, string> f2vs = new Dictionary<string, string>();
                    f2vs.Add("detailPageUrl", priceUrl);
                    f2vs.Add("detailPageName", pageName);
                    f2vs.Add("cookie", cookie);
                    f2vs.Add("category1Code", category1Code);
                    f2vs.Add("category2Code", category2Code);
                    f2vs.Add("category1Name", category1Name);
                    f2vs.Add("category2Name", category2Name);
                    f2vs.Add("district", district);
                    f2vs.Add("productSysNo", productSysNo);
                    f2vs.Add("商品名称", name);
                    f2vs.Add("品牌", pinpai);
                    f2vs.Add("净含量", jinghanliang);
                    f2vs.Add("产品毛重", chanpinmaozhong);
                    f2vs.Add("产地", chandi);
                    f2vs.Add("保质期", baozhiqi);
                    f2vs.Add("规格", guige);
                    f2vs.Add("商品页URL", pageUrl);
                    f2vs.Add("营销栏目", yingxiaolanmu);
                    resultEW.AddRow(f2vs);
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

        private bool GenerateRedirectPageInfo(IListSheet listSheet)
        {
            bool succeed = true;
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            string exportDir = this.RunPage.GetExportDir();
            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("productSysNo", 5);
            resultColumnDic.Add("category1Code", 6);
            resultColumnDic.Add("category2Code", 7);
            resultColumnDic.Add("category1Name", 8);
            resultColumnDic.Add("category2Name", 9);
            resultColumnDic.Add("district", 10);
            string resultFilePath = Path.Combine(exportDir, "我买网获取所有详情页.xlsx");

            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic);

            GetRedirectPageUrl(listSheet, pageSourceDir, resultEW);

            resultEW.SaveToDisk();

            return succeed;
        }

        private void GetRedirectPageUrl(IListSheet listSheet, string pageSourceDir, ExcelWriter resultEW)
        {
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string pageUrl = listSheet.PageUrlList[i];
                string pageName = listSheet.PageNameList[i];
                string localFilePath = this.RunPage.GetFilePath(pageUrl, pageSourceDir);
                string productSysNo = row["productSysNo"];
                string category1Code = row["category1Code"];
                string category2Code = row["category2Code"];
                string category1Name = row["category1Name"];
                string category2Name = row["category2Name"];
                string district = row["district"];
                string cookie = row["cookie"];

                TextReader tr = null;

                try
                {
                    tr = new StreamReader(localFilePath);
                    string webPageHtml = tr.ReadToEnd().Trim();
                    if (webPageHtml.StartsWith("<script>location"))
                    {
                        string[] strs = webPageHtml.Split(new string[] { "\"" }, StringSplitOptions.RemoveEmptyEntries);
                        string urlStr = strs[1];
                        if (urlStr.StartsWith("http://"))
                        {
                            pageUrl = urlStr;
                        }
                        else
                        {
                            pageUrl = "http://www.womai.com" + urlStr;
                        }
                    }


                    Dictionary<string, string> f2vs = new Dictionary<string, string>();
                    f2vs.Add("detailPageUrl", pageUrl);
                    f2vs.Add("detailPageName", pageName);
                    f2vs.Add("cookie", cookie);
                    f2vs.Add("category1Code", category1Code);
                    f2vs.Add("category2Code", category2Code);
                    f2vs.Add("category1Name", category1Name);
                    f2vs.Add("category2Name", category2Name);
                    f2vs.Add("district", district);
                    f2vs.Add("productSysNo", productSysNo);
                    resultEW.AddRow(f2vs);
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
    }
}