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
using System.Net;

namespace NetDataAccess.Extended.Fortune
{
    public class GetCNNCompanyDetailPage : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            this.GetInfos(listSheet);

            return true;
        }

        public override bool AfterGrabOneCatchException(string pageUrl, Dictionary<string, string> listRow, Exception ex)
        {
            if (ex.InnerException is WebException)
            {
                WebException webEx = (WebException)ex.InnerException;
                if (webEx.Status == WebExceptionStatus.ProtocolError)
                {
                    HttpWebResponse webRes = (HttpWebResponse)webEx.Response;
                    if (webRes.StatusCode == HttpStatusCode.NotFound)
                    {
                        this.RunPage.InvokeAppendLogText("服务器端不存在此网页(404), pageUrl = " + pageUrl, LogLevelType.Error, true);
                        return true;
                    }
                }
            }
            return false;
        }

        private string GetNextText(HtmlNode titleNode)
        {
            HtmlNode nextNode = titleNode.NextSibling;
            while (nextNode != null)
            {
                if (nextNode.Name.ToLower() == "br")
                {
                    nextNode = nextNode.NextSibling;
                }
                else if (nextNode.Name.ToLower() == "text")
                {
                    return CommonUtil.HtmlDecode(nextNode.InnerText).Trim();
                }
            }
            return "";
        }

        private void GetInfos(IListSheet listSheet)
        {
            string sourceDir = this.RunPage.GetDetailSourceFileDir();

            ExcelWriter resultEW = this.CreateCompanyInfoWriter();

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> listRow = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(listRow[SysConfig.GiveUpGrabFieldName]);
                string companyPageUrl = listRow[SysConfig.DetailPageUrlFieldName];
                string year = listRow["yearValue"];
                string companyName = listRow["companyName"];
                string companyID = listRow["companyID"];
                string industryName = listRow["industryName"];
                string score = listRow["score"];

                Dictionary<string, string> resultRow = new Dictionary<string, string>();
                resultRow.Add("url", companyPageUrl);
                resultRow.Add("year", year);
                resultRow.Add("companyName", companyName);
                resultRow.Add("industryName", industryName);
                resultRow.Add("score", score);
                string innovation = "";
                string peopleManagement = "";
                string useOfCorporateAssets = "";
                string socialResponsibility = "";
                string qualityOfManagement = "";
                string financialSoundness = "";
                string longTermInvestment = "";
                string qualityOfProductsOrServices = "";
                string globalCompetitiveness = "";
                if (!giveUp)
                {
                    try
                    {

                        string localFilePath = this.RunPage.GetFilePath(companyPageUrl, sourceDir);
                        string fileText = FileHelper.GetTextFromFile(localFilePath, Encoding.UTF8);
                        switch (year)
                        {
                            case "2013":
                            case "2014":
                                {
                                    HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                                    htmlDoc.LoadHtml(fileText);
                                    HtmlNodeCollection rankInfoNodes = htmlDoc.DocumentNode.SelectNodes("//div[@class=\"snapDataRow\"]");
                                    if(rankInfoNodes!=null){
                                        foreach (HtmlNode rankInfoNode in rankInfoNodes)
                                        {
                                            string rankName = CommonUtil.HtmlDecode(rankInfoNode.SelectSingleNode("./div[@class=\"cnncol1\"]").InnerText).Trim();
                                            string rankValue = CommonUtil.HtmlDecode(rankInfoNode.SelectSingleNode("./div[@class=\"cnncol2\"]").InnerText).Trim();
                                            switch (rankName)
                                            {
                                                case "Innovation":
                                                    innovation = rankValue;
                                                    break;
                                                case "People management":
                                                    peopleManagement = rankValue;
                                                    break;
                                                case "Use of corporate assets":
                                                    useOfCorporateAssets = rankValue;
                                                    break;
                                                case "Social responsibility":
                                                    socialResponsibility = rankValue;
                                                    break;
                                                case "Quality of management":
                                                    qualityOfManagement = rankValue;
                                                    break;
                                                case "Financial soundness":
                                                    financialSoundness = rankValue;
                                                    break;
                                                case "Long-term investment":
                                                    longTermInvestment = rankValue;
                                                    break;
                                                case "Quality of products/services":
                                                    qualityOfProductsOrServices = rankValue;
                                                    break;
                                                case "Global competitiveness":
                                                    globalCompetitiveness = rankValue;
                                                    break;
                                            }
                                        }
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

                                    HtmlNodeCollection rankLineNodes = htmlDoc.DocumentNode.SelectNodes("//table[@class=\"cnnwith220inset\"]/tr");
                                    foreach (HtmlNode rankLineNode in rankLineNodes)
                                    {
                                        HtmlNodeCollection tdNodes = rankLineNode.SelectNodes("./td");
                                        if (tdNodes != null)
                                        {
                                            string rankName = CommonUtil.HtmlDecode(rankLineNode.SelectSingleNode("./td[@class=\"cnncol1\"]/a").InnerText).Trim();
                                            string rankValue = CommonUtil.HtmlDecode(rankLineNode.SelectSingleNode("./td[@class=\"cnncol2\"]").InnerText).Trim();
                                            switch (rankName)
                                            {
                                                case "Innovation":
                                                    innovation = rankValue;
                                                    break;
                                                case "People management":
                                                    peopleManagement = rankValue;
                                                    break;
                                                case "Use of corporate assets":
                                                    useOfCorporateAssets = rankValue;
                                                    break;
                                                case "Social responsibility":
                                                    socialResponsibility = rankValue;
                                                    break;
                                                case "Quality of management":
                                                    qualityOfManagement = rankValue;
                                                    break;
                                                case "Financial soundness":
                                                    financialSoundness = rankValue;
                                                    break;
                                                case "Long-term investment":
                                                    longTermInvestment = rankValue;
                                                    break;
                                                case "Quality of products/services":
                                                    qualityOfProductsOrServices = rankValue;
                                                    break;
                                                case "Global competitiveness":
                                                    globalCompetitiveness = rankValue;
                                                    break;
                                            }
                                        }
                                    }
                                }
                                break;
                            case "2007":
                            case "2006":
                                {
                                    HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                                    htmlDoc.LoadHtml(fileText);

                                    HtmlNodeCollection rankLineNodes = htmlDoc.DocumentNode.SelectNodes("//div/div/table[@class=\"maglisttable\"]/tr[@id=\"tablerow\"]");
                                    foreach (HtmlNode rankLineNode in rankLineNodes)
                                    {
                                        string rankName = CommonUtil.HtmlDecode(rankLineNode.SelectSingleNode("./td[@class=\"textcell\"]/a").InnerText).Trim();
                                        string rankValue = CommonUtil.HtmlDecode(rankLineNode.SelectSingleNode("./td[@class=\"datacell\"]").InnerText).Trim();
                                        switch (rankName)
                                        {
                                            case "Innovation":
                                                innovation = rankValue;
                                                break;
                                            case "People management":
                                                peopleManagement = rankValue;
                                                break;
                                            case "Use of corporate assets":
                                                useOfCorporateAssets = rankValue;
                                                break;
                                            case "Social responsibility":
                                                socialResponsibility = rankValue;
                                                break;
                                            case "Quality of management":
                                                qualityOfManagement = rankValue;
                                                break;
                                            case "Financial soundness":
                                                financialSoundness = rankValue;
                                                break;
                                            case "Long-term investment":
                                                longTermInvestment = rankValue;
                                                break;
                                            case "Quality of products/services":
                                                qualityOfProductsOrServices = rankValue;
                                                break;
                                            case "Global competitiveness":
                                                globalCompetitiveness = rankValue;
                                                break;
                                        }
                                    }
                                }
                                break;
                            case "2008":
                                {
                                    HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                                    htmlDoc.LoadHtml(fileText);

                                    HtmlNodeCollection rankLineNodes = htmlDoc.DocumentNode.SelectNodes("//div[@id=\"magFeatData\"]/table/tr[@id=\"tablerow\"]");
                                    foreach (HtmlNode rankLineNode in rankLineNodes)
                                    {
                                        string rankName = CommonUtil.HtmlDecode(rankLineNode.SelectSingleNode("./td[@class=\"alignLft\"]/a").InnerText).Trim();
                                        string rankValue = CommonUtil.HtmlDecode(rankLineNode.SelectSingleNode("./td[@class=\"alignRgt\"]").InnerText).Trim();
                                        switch (rankName)
                                        {
                                            case "Innovation":
                                                innovation = rankValue;
                                                break;
                                            case "People management":
                                                peopleManagement = rankValue;
                                                break;
                                            case "Use of corporate assets":
                                                useOfCorporateAssets = rankValue;
                                                break;
                                            case "Social responsibility":
                                                socialResponsibility = rankValue;
                                                break;
                                            case "Quality of management":
                                                qualityOfManagement = rankValue;
                                                break;
                                            case "Financial soundness":
                                                financialSoundness = rankValue;
                                                break;
                                            case "Long-term investment":
                                                longTermInvestment = rankValue;
                                                break;
                                            case "Quality of products/services":
                                                qualityOfProductsOrServices = rankValue;
                                                break;
                                            case "Global competitiveness":
                                                globalCompetitiveness = rankValue;
                                                break;
                                        }
                                    }
                                }
                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        this.RunPage.InvokeAppendLogText(ex.Message + ". 解析出错， pageUrl = " + companyPageUrl, LogLevelType.Error, true);
                        throw ex;
                    }
                }

                resultRow.Add("innovation", innovation);
                resultRow.Add("peopleManagement", peopleManagement);
                resultRow.Add("useOfCorporateAssets", useOfCorporateAssets);
                resultRow.Add("socialResponsibility", socialResponsibility);
                resultRow.Add("qualityOfManagement", qualityOfManagement);
                resultRow.Add("financialSoundness", financialSoundness);
                resultRow.Add("longTermInvestment", longTermInvestment);
                resultRow.Add("qualityOfProductsOrServices", qualityOfProductsOrServices);
                resultRow.Add("globalCompetitiveness", globalCompetitiveness);
                resultEW.AddRow(resultRow);
            }

            resultEW.SaveToDisk();
        }

        private ExcelWriter CreateCompanyInfoWriter()
        {
            String exportDir = this.RunPage.GetExportDir();
            string resultFilePath = Path.Combine(exportDir, "Fortune_CNN_CompanyList.xlsx");

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("url", 0);
            resultColumnDic.Add("year", 1);
            resultColumnDic.Add("companyName", 2);
            resultColumnDic.Add("companyID", 3);
            resultColumnDic.Add("industryName", 4);
            resultColumnDic.Add("industryRank", 5);
            resultColumnDic.Add("score", 6);

            resultColumnDic.Add("innovation", 7);
            resultColumnDic.Add("peopleManagement", 8);
            resultColumnDic.Add("useOfCorporateAssets", 9);
            resultColumnDic.Add("socialResponsibility", 10);
            resultColumnDic.Add("qualityOfManagement", 11);
            resultColumnDic.Add("financialSoundness", 12);
            resultColumnDic.Add("longTermInvestment", 13);
            resultColumnDic.Add("qualityOfProductsOrServices", 14);
            resultColumnDic.Add("globalCompetitiveness", 15);
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }
    }
}