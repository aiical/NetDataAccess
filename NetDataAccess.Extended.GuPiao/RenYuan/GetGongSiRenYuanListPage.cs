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

namespace NetDataAccess.Extended.GuPiao
{
    public class GetGongSiRenYuanListPage : ExternalRunWebPage
    { 
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            //this.GetRenYuanListInfos(listSheet);
            this.GetRenYuanDetailPageUrls(listSheet);
            return true;
        }

        private void GetRenYuanListInfos(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("公司代码", 0);
            resultColumnDic.Add("姓名", 1);
            resultColumnDic.Add("职务", 2);
            resultColumnDic.Add("起始日期", 3);
            resultColumnDic.Add("终止日期", 4);
            string resultFilePath = Path.Combine(exportDir, "上市公司高管人员列表.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"];
                string code = row["code"];

                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    HtmlAgilityPack.HtmlDocument pageHtmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i, Encoding.GetEncoding("gb2312"));

                    try
                    {
                        HtmlNodeCollection tableList = pageHtmlDoc.DocumentNode.SelectNodes("//table[@id=\"comInfo1\"]");

                        foreach (HtmlNode tableNode in tableList)
                        {
                            HtmlNode categoryNode = tableNode.SelectSingleNode("./thead");
                            string categoryName = categoryNode.InnerText.Trim();
                            string periodName = "";
                            HtmlNodeCollection itemTrNodeList = tableNode.SelectNodes("./tbody/tr");
                            for (int j = 1; j < itemTrNodeList.Count; j++)
                            {
                                HtmlNode itemTrNode = itemTrNodeList[j];
                                HtmlNodeCollection itemTdNodesList = itemTrNode.SelectNodes("./td");
                                if (itemTdNodesList.Count == 1)
                                {
                                    periodName = itemTdNodesList[0].InnerText.Trim();
                                }
                                else
                                {
                                    Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                    f2vs.Add("公司代码", code);
                                    f2vs.Add("姓名", itemTdNodesList[0].InnerText.Trim());
                                    f2vs.Add("职务", itemTdNodesList[1].InnerText.Trim());
                                    f2vs.Add("起始日期", itemTdNodesList[2].InnerText.Trim());
                                    f2vs.Add("终止日期", itemTdNodesList[3].InnerText.Trim());
                                    resultEW.AddRow(f2vs);
                                }
                            }
                        }

                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
            }
            resultEW.SaveToDisk();
        }


        private void GetRenYuanDetailPageUrls(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("公司代码", 5);
            resultColumnDic.Add("姓名", 6);
            string resultFilePath = Path.Combine(exportDir, "上市公司高管人员详情页.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);

            Dictionary<string, string> urlDic = new Dictionary<string, string>();
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"];
                string code = row["code"];

                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    HtmlAgilityPack.HtmlDocument pageHtmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i, Encoding.GetEncoding("gb2312"));

                    try
                    {
                        HtmlNodeCollection tableList = pageHtmlDoc.DocumentNode.SelectNodes("//table[@id=\"comInfo1\"]");

                        foreach (HtmlNode tableNode in tableList)
                        {
                            HtmlNode categoryNode = tableNode.SelectSingleNode("./thead");
                            string categoryName = categoryNode.InnerText.Trim();
                            string periodName = "";
                            HtmlNodeCollection itemTrNodeList = tableNode.SelectNodes("./tbody/tr");
                            for (int j = 1; j < itemTrNodeList.Count; j++)
                            {
                                HtmlNode itemTrNode = itemTrNodeList[j];
                                HtmlNodeCollection itemTdNodesList = itemTrNode.SelectNodes("./td");
                                if (itemTdNodesList.Count == 1)
                                {
                                    periodName = itemTdNodesList[0].InnerText.Trim();
                                }
                                else
                                {
                                    HtmlNode linkNode = itemTdNodesList[0].SelectSingleNode("./div/a");
                                    string url = "http://vip.stock.finance.sina.com.cn" + linkNode.GetAttributeValue("href", "");
                                    if (!urlDic.ContainsKey(url))
                                    {
                                        urlDic.Add(url, null);

                                        int urlPostfixIndex = url.LastIndexOf("Name=");
                                        string urlPostfix = url.Substring(urlPostfixIndex + 5);
                                        url = url.Substring(0, urlPostfixIndex + 5) + CommonUtil.StringToHexString(urlPostfix, Encoding.GetEncoding("gb2312"));
                                        
                                        Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                        f2vs.Add("detailPageUrl", url);
                                        f2vs.Add("detailPageName", url);
                                        f2vs.Add("公司代码", code);
                                        f2vs.Add("姓名", itemTdNodesList[0].InnerText.Trim());

                                        resultEW.AddRow(f2vs);
                                    }
                                }
                            }
                        }

                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
            }
            resultEW.SaveToDisk();
        }
         
    }
}