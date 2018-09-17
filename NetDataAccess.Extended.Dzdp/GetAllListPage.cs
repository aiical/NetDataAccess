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

namespace NetDataAccess.Extended.Dzdp
{
    public class GetAllListPage : ExternalRunWebPage
    {
        public override void CheckRequestCompleteFile(string webPageText, Dictionary<string, string> listRow)
        {
            if (webPageText.Contains("关于大众点评") && webPageText.Trim().EndsWith("</html>"))
            {
            }
            else
            {
                throw new Exception("未完全加载文件.");
            }
        }

        public ExcelWriter GetExcelWriter(int fileIndex)
        {
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("city", 5);
            resultColumnDic.Add("g", 6);
            resultColumnDic.Add("r", 7);
            resultColumnDic.Add("gName", 8);
            resultColumnDic.Add("rName", 9);
            resultColumnDic.Add("pageIndex", 10);
            resultColumnDic.Add("shopName", 11);
            resultColumnDic.Add("shopCode", 12);
            resultColumnDic.Add("reviewNum", 13);
            resultColumnDic.Add("serviceRating", 14);
            resultColumnDic.Add("environmentRating", 15);
            resultColumnDic.Add("tasteRating", 16);
            resultColumnDic.Add("address", 17);

            string resultFilePath = Path.Combine(exportDir, "大众点评获取店铺详情_" + fileIndex.ToString() + ".xlsx");
            Dictionary<string, string> resultColumnFormat = new Dictionary<string, string>();
            resultColumnFormat.Add("reviewNum", "#,##0");
            resultColumnFormat.Add("serviceRating", "#0.0");
            resultColumnFormat.Add("environmentRating", "#0.0");
            resultColumnFormat.Add("tasteRating", "#0.0");

            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, resultColumnFormat);
            return resultEW;
        }

        public override bool AfterAllGrab(IListSheet listSheet)
        {
            ExcelWriter resultEW = null;
            int fileIndex = 1;

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;

            Dictionary<string, string> shopDic = new Dictionary<string, string>();

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                if (resultEW == null || resultEW.RowCount > 500000)
                {
                    if (resultEW != null)
                    {
                        resultEW.SaveToDisk();
                    }
                    resultEW = this.GetExcelWriter(fileIndex);
                    fileIndex++;
                }

                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string url = row[detailPageUrlColumnName];
                    string cookie = row["cookie"];
                    string city = row["city"];
                    string g = row["g"];
                    string r = row["r"];
                    string gName = row["gName"];
                    string rName = row["rName"];
                    string pageIndex = row["pageIndex"]; 

                    HtmlAgilityPack.HtmlDocument pageHtmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                    HtmlNodeCollection allShopLinkNodes = pageHtmlDoc.DocumentNode.SelectNodes("//div[@id=\"shop-all-list\"]/ul/li/div[@class=\"txt\"]");

                    if(allShopLinkNodes!=null){
                        for (int j = 0; j < allShopLinkNodes.Count; j++)
                        {
                            HtmlNode shopLinkNode = allShopLinkNodes[j];
                            HtmlNodeCollection ratingNodes = shopLinkNode.SelectNodes("./span[@class=\"comment-list\"]/span");
                            Nullable<decimal> serviceRating = null;
                            Nullable<decimal> environmentRating = null;
                            Nullable<decimal> tasteRating = null;

                            if (ratingNodes != null)
                            {
                                foreach (HtmlNode ratingNode in ratingNodes)
                                {
                                    string ratingText = ratingNode.InnerText.Trim();
                                    decimal ratingValue = decimal.Parse(ratingNode.SelectSingleNode("./b").InnerText.Trim());
                                    if (ratingText.StartsWith("口味"))
                                    {
                                        tasteRating = ratingValue;
                                    }
                                    else if (ratingText.StartsWith("服务"))
                                    {
                                        serviceRating = ratingValue;
                                    }
                                    else if (ratingText.StartsWith("环境"))
                                    {
                                        environmentRating = ratingValue;
                                    }
                                }
                            }

                            string address = "";
                            HtmlNode addressNode = shopLinkNode.SelectSingleNode("./div[@class=\"tag-addr\"]/span[@class=\"addr\"]");
                            if (addressNode != null)
                            {
                                address = CommonUtil.HtmlDecode(addressNode.InnerText.Trim());
                            }

                            HtmlNode nameNode = shopLinkNode.SelectSingleNode("./div[@class=\"tit\"]/a");
                            HtmlNode reviewNumNode = shopLinkNode.SelectSingleNode("./div[@class=\"comment\"]/a[@class=\"review-num\"]/b");
                            if (nameNode != null)
                            {
                                string shopName = nameNode.Attributes["title"].Value;
                                string hrefValue = nameNode.Attributes["href"].Value;
                                string[] hrefs = hrefValue.Split(new string[] { "/" }, StringSplitOptions.RemoveEmptyEntries);
                                string shopCode = hrefs[hrefs.Length - 1];

                                Nullable<int> reviewNum = 0;
                                if (reviewNumNode != null)
                                {
                                    reviewNum = int.Parse(reviewNumNode.InnerText);
                                }

                                if (!shopDic.ContainsKey(shopCode))
                                {
                                    shopDic.Add(shopCode, "");

                                    string detailPageName = shopCode;
                                    string detailPageUrl = hrefValue;
                                    Dictionary<string, object> f2vs = new Dictionary<string, object>();
                                    f2vs.Add("detailPageUrl", detailPageUrl);
                                    f2vs.Add("detailPageName", detailPageName);
                                    f2vs.Add("cookie", "cy=22; cye=jinan;");
                                    f2vs.Add("city", city);
                                    f2vs.Add("g", g);
                                    f2vs.Add("r", r);
                                    f2vs.Add("gName", gName);
                                    f2vs.Add("rName", rName);
                                    f2vs.Add("pageIndex", pageIndex);
                                    f2vs.Add("shopName", shopName);
                                    f2vs.Add("shopCode", shopCode);
                                    f2vs.Add("reviewNum", reviewNum);
                                    f2vs.Add("serviceRating", serviceRating);
                                    f2vs.Add("environmentRating", environmentRating);
                                    f2vs.Add("tasteRating", tasteRating);
                                    f2vs.Add("address", address);
                                    resultEW.AddRow(f2vs);
                                }
                            }
                        }
                    }
                }
            }

            resultEW.SaveToDisk();

            return true;
        } 
    }
}