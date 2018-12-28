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

namespace NetDataAccess.Extended.Fortune
{
    public class GetCNNYearListPage : ExternalRunWebPage
    { 
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            this.GetInfos(listSheet); 

            return true;
        }



        private void GetInfos(IListSheet listSheet)
        {
            string sourceDir = this.RunPage.GetDetailSourceFileDir();
             
            ExcelWriter resultEW =  this.CreateCompanyPageWriter(); 

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> listRow = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(listRow[SysConfig.GiveUpGrabFieldName]);
                string listPageUrl = listRow[SysConfig.DetailPageUrlFieldName];
                string year = listRow["yearValue"];
                if (!giveUp)
                {
                    try
                    {
                        string localFilePath = this.RunPage.GetFilePath(listPageUrl, sourceDir);
                        string fileText = FileHelper.GetTextFromFile(localFilePath, Encoding.UTF8);
                        switch (year)
                        {
                            case "2013":
                            case "2014":
                                {
                                    JArray fileJArray = JArray.Parse(fileText);
                                    for (int j = 0; j < fileJArray.Count; j++)
                                    {
                                        JObject companObj = fileJArray[j] as JObject;
                                        string companyName = companObj.GetValue("companyName").ToString();
                                        string industryName = companObj.GetValue("industryName").ToString();
                                        string state = companObj.GetValue("state").ToString();
                                        string score = companObj.GetValue("score").ToString();
                                        string top50rank = companObj.GetValue("top50rank").ToString();
                                        string industryRank = companObj.GetValue("industryRank").ToString();
                                        string previousRank = companObj.GetValue("previousRank").ToString();
                                        string companyID = companObj.GetValue("companyID").ToString();
                                        string inTop50 = companObj.GetValue("inTop50").ToString();
                                        string location = companObj.GetValue("location").ToString();

                                        string companyPageUrl = "https://money.cnn.com/magazines/fortune/most-admired/" + year + "/snapshots/" + companyID + ".html?iid=wma14_fl_list";

                                        Dictionary<string, string> resultRow = new Dictionary<string, string>();
                                        resultRow.Add("detailPageUrl", companyPageUrl);
                                        resultRow.Add("detailPageName", companyPageUrl);
                                        resultRow.Add("yearValue", year);
                                        resultRow.Add("companyName", companyName);
                                        resultRow.Add("industryName", industryName);
                                        resultRow.Add("state", state);
                                        resultRow.Add("score", score);
                                        resultRow.Add("top50rank", top50rank);
                                        resultRow.Add("industryRank", industryRank);
                                        resultRow.Add("previousRank", previousRank);
                                        resultRow.Add("companyID", companyID);
                                        resultRow.Add("inTop50", inTop50);
                                        resultRow.Add("location", location);                                        
                                        resultEW.AddRow(resultRow);
                                    }
                                }
                                break;
                            case "2012":
                            case "2011":
                            case "2010":
                            case "2009": 
                                {
                                    HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                                    htmlDoc.LoadHtml(fileText);

                                    int dirEndIndex = listPageUrl.IndexOf("/companies");
                                    string urlDir = listPageUrl.Substring(0, dirEndIndex);

                                    HtmlNodeCollection companyLineNodes = htmlDoc.DocumentNode.SelectNodes("//table[@class=\"cnnwith220inset\"]/tbody/tr");
                                    foreach (HtmlNode companyLineNode in companyLineNodes)
                                    {
                                        HtmlNodeCollection companyInfoNodes = companyLineNode.SelectNodes("./td");
                                        HtmlNode nameNode = companyInfoNodes[0].SelectSingleNode("./a");
                                        string companyPageUrl = urlDir + nameNode.GetAttributeValue("href", "").Substring(2);
                                        string companyName = CommonUtil.HtmlDecode(nameNode.InnerText).Trim();
                                        string industryName = CommonUtil.HtmlDecode(companyInfoNodes[1].InnerText).Trim();
                                        string score = CommonUtil.HtmlDecode(companyInfoNodes[2].InnerText).Trim();

                                        Dictionary<string, string> resultRow = new Dictionary<string, string>();
                                        resultRow.Add("detailPageUrl", companyPageUrl);
                                        resultRow.Add("detailPageName", companyPageUrl);
                                        resultRow.Add("yearValue", year);
                                        resultRow.Add("companyName", companyName);
                                        resultRow.Add("industryName", industryName);
                                        resultRow.Add("score", score);
                                        resultEW.AddRow(resultRow);
                                    }
                                }
                                break;
                            case "2007":
                            case "2006":
                                {
                                    HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                                    htmlDoc.LoadHtml(fileText);

                                    int dirEndIndex = listPageUrl.IndexOf("/companies");
                                    string urlDir = listPageUrl.Substring(0, dirEndIndex);

                                    HtmlNodeCollection companyLineNodes = htmlDoc.DocumentNode.SelectNodes("//table[@class=\"maglisttable\"]/tr[@id=\"tablerow\"]");
                                    foreach (HtmlNode companyLineNode in companyLineNodes)
                                    {
                                        HtmlNodeCollection companyInfoNodes = companyLineNode.SelectNodes("./td");
                                        HtmlNode nameNode = companyInfoNodes[0].SelectSingleNode("./a");
                                        string companyPageUrl = urlDir + nameNode.GetAttributeValue("href", "").Substring(2);
                                        string companyName = CommonUtil.HtmlDecode(nameNode.InnerText).Trim();
                                        string industryName = CommonUtil.HtmlDecode(companyInfoNodes[1].InnerText).Trim();
                                        string score = CommonUtil.HtmlDecode(companyInfoNodes[2].InnerText).Trim();

                                        Dictionary<string, string> resultRow = new Dictionary<string, string>();
                                        resultRow.Add("detailPageUrl", companyPageUrl);
                                        resultRow.Add("detailPageName", companyPageUrl);
                                        resultRow.Add("yearValue", year);
                                        resultRow.Add("companyName", companyName);
                                        resultRow.Add("industryName", industryName);
                                        resultRow.Add("score", score);
                                        resultEW.AddRow(resultRow);
                                    }
                                }
                                break;
                            case "2008":
                                {
                                    HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                                    htmlDoc.LoadHtml(fileText);

                                    int dirEndIndex = listPageUrl.IndexOf("/companies");
                                    string urlDir = listPageUrl.Substring(0, dirEndIndex);

                                    HtmlNodeCollection companyLineNodes = htmlDoc.DocumentNode.SelectNodes("//div[@id=\"magFeatData\"]/table/tr[@id=\"tablerow\"]");
                                    foreach (HtmlNode companyLineNode in companyLineNodes)
                                    {
                                        HtmlNodeCollection companyInfoNodes = companyLineNode.SelectNodes("./td");
                                        HtmlNode nameNode = companyInfoNodes[0].SelectSingleNode("./a");
                                        string companyPageUrl = urlDir + nameNode.GetAttributeValue("href", "").Substring(2);
                                        string companyName = CommonUtil.HtmlDecode(nameNode.InnerText).Trim();
                                        string industryName = CommonUtil.HtmlDecode(companyInfoNodes[1].InnerText).Trim();
                                        string score = CommonUtil.HtmlDecode(companyInfoNodes[2].InnerText).Trim();

                                        Dictionary<string, string> resultRow = new Dictionary<string, string>();
                                        resultRow.Add("detailPageUrl", companyPageUrl);
                                        resultRow.Add("detailPageName", companyPageUrl);
                                        resultRow.Add("yearValue", year);
                                        resultRow.Add("companyName", companyName);
                                        resultRow.Add("industryName", industryName);
                                        resultRow.Add("score", score);
                                        resultEW.AddRow(resultRow);
                                    }
                                }
                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        this.RunPage.InvokeAppendLogText(ex.Message + ". 解析出错， pageUrl = " + listPageUrl, LogLevelType.Error, true);
                        throw ex;
                    }
                }
            }

            resultEW.SaveToDisk();
        }         

        private ExcelWriter CreateCompanyPageWriter()
        {
            String exportDir = this.RunPage.GetExportDir();
            string resultFilePath = Path.Combine(exportDir, "Fortune_CNN_CompanyDetail.xlsx");

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("yearValue", 5);
            resultColumnDic.Add("companyName", 6);
            resultColumnDic.Add("industryName", 7);
            resultColumnDic.Add("state", 8);
            resultColumnDic.Add("score", 9);
            resultColumnDic.Add("top50rank", 10);
            resultColumnDic.Add("industryRank", 11);
            resultColumnDic.Add("previousRank", 12);
            resultColumnDic.Add("companyID", 13);
            resultColumnDic.Add("inTop50", 14);
            resultColumnDic.Add("location", 15); 
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }         
    }
}