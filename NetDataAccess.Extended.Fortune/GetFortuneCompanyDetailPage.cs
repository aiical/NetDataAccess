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
    public class GetFortuneCompanyDetailPage : ExternalRunWebPage
    { 
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            this.GetInfos(listSheet); 

            return true;
        }

        public override void WebRequestHtml_BeforeSendRequest(string pageUrl, Dictionary<string, string> listRow, Base.Web.NDAWebClient client)
        {
            client.Headers.Add("accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8");
            client.Headers.Add("Upgrade-Insecure-Requests", "1");
            client.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/67.0.3396.99 Safari/537.36");
            client.Headers.Add("Cache-Control", "max-age=0");
            client.Headers.Add("Host", "fortune.com");
        }

        public override bool AfterGrabOneCatchException(string pageUrl, Dictionary<string, string> listRow, Exception ex)
        {
            return base.AfterGrabOneCatchException(pageUrl, listRow, ex);
        }

        private void GetInfos(IListSheet listSheet)
        {
            string sourceDir = this.RunPage.GetDetailSourceFileDir();
             
            ExcelWriter resultEW =  this.CreateCompanyPageWriter();

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> listRow = listSheet.GetRow(i);
                string companyPageUrl = listRow[SysConfig.DetailPageUrlFieldName];
                string year = listRow["yearValue"];
                try
                {
                    string localFilePath = this.RunPage.GetFilePath(companyPageUrl, sourceDir);
                    switch (year)
                    {
                        case "2018":
                        case "2017": 
                            {
                                string fileText = FileHelper.GetTextFromFile(localFilePath, Encoding.UTF8);
                                string industry = "";
                                string industryRank = "";
                                string score = "";

                                HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                                htmlDoc.LoadHtml(fileText);
                                //HtmlNodeCollection baseInfoNodes = htmlDoc.DocumentNode.SelectNodes("//div[@class=\"company-info-card\"]/div[@class=\"row expanded\"]/div[@class=\"small-12 large-7 columns bg-white company-info-c\"]/div[@class=\"row company-info-card-table\"]/div[@class=\"row\"]");
                                HtmlNodeCollection baseInfoNodes = htmlDoc.DocumentNode.SelectNodes("//div/div/div/div/div[@class=\"row\"]");
                                foreach (HtmlNode baseInfoNode in baseInfoNodes)
                                {
                                    HtmlNode infoTitleNode = baseInfoNode.SelectSingleNode("./p");
                                    HtmlNode infoValueNode = baseInfoNode.SelectSingleNode("./div/p");
                                    if (infoTitleNode != null && infoValueNode != null)
                                    {
                                        string infoTitle = CommonUtil.HtmlDecode(infoTitleNode.InnerText).Trim();
                                        string infoValue = CommonUtil.HtmlDecode(infoValueNode.InnerText).Trim();
                                        switch (infoTitle)
                                        {
                                            case "Industry":
                                                industry = infoValue;
                                                break;
                                            case "Industry Ranking":
                                                industryRank = infoValue;
                                                break;
                                            case "Overall Score":
                                                score = infoValue;
                                                break;
                                        }
                                    }
                                }
                                string financialSoundness = "";
                                string globalCompetitiveness = "";
                                string innovation = "";
                                string longTermInvestmentValue = "";
                                string peopleManagement = "";
                                string qualityOfManagement = "";
                                string qualityOfProductsOrServices = "";
                                string socialResponsibility = "";
                                string useOfCorporateAssets = "";

                                HtmlNodeCollection externalInfoNodes = htmlDoc.DocumentNode.SelectNodes("//div[@class=\"row expanded bg-grey\"]/div/div/div/div/div/div/div/table/tbody/tr");
                                foreach (HtmlNode externalInfoNode in externalInfoNodes)
                                {
                                    HtmlNodeCollection tdNodes = externalInfoNode.SelectNodes("./td");
                                    string rankTitle = CommonUtil.HtmlDecode(tdNodes[0].InnerText).Trim();
                                    string rankValue = CommonUtil.HtmlDecode(tdNodes[1].InnerText).Trim();
                                    switch (rankTitle)
                                    {
                                        case "Financial Soundness":
                                            financialSoundness = rankValue;
                                            break;
                                        case "Global Competitiveness":
                                            globalCompetitiveness = rankValue;
                                            break;
                                        case "Innovation":
                                            innovation = rankValue;
                                            break;
                                        case "Long-Term Investment Value":
                                            longTermInvestmentValue = rankValue;
                                            break;
                                        case "People Management":
                                            peopleManagement = rankValue;
                                            break;
                                        case "Quality of Management":
                                            qualityOfManagement = rankValue;
                                            break;
                                        case "Quality of Products/Services":
                                            qualityOfProductsOrServices = rankValue;
                                            break;
                                        case "Social Responsibility":
                                            socialResponsibility = rankValue;
                                            break;
                                        case "Use of Corporate Assets":
                                            useOfCorporateAssets = rankValue;
                                            break;
                                    }
                                }

                                Dictionary<string, string> resultRow = new Dictionary<string, string>();
                                resultRow.Add("url", companyPageUrl);
                                resultRow.Add("companyName", listRow["companyName"]);
                                resultRow.Add("companyID", listRow["companyID"]);
                                resultRow.Add("year", listRow["yearValue"]);

                                resultRow.Add("industryName", industry);
                                resultRow.Add("industryRank", industryRank);
                                resultRow.Add("score", score);

                                resultRow.Add("financialSoundness", financialSoundness);
                                resultRow.Add("globalCompetitiveness", globalCompetitiveness);
                                resultRow.Add("innovation", innovation);
                                resultRow.Add("longTermInvestmentValue", longTermInvestmentValue);
                                resultRow.Add("peopleManagement", peopleManagement);
                                resultRow.Add("qualityOfManagement", qualityOfManagement);
                                resultRow.Add("qualityOfProductsOrServices", qualityOfProductsOrServices);
                                resultRow.Add("socialResponsibility", socialResponsibility);
                                resultRow.Add("useOfCorporateAssets", useOfCorporateAssets);

                                resultEW.AddRow(resultRow);
                            }
                            break; 
                        case "2016":
                            {
                                string fileText = FileHelper.GetTextFromFile(localFilePath, Encoding.UTF8);
                                string industry = "";
                                string industryRank = "";
                                string score = "";
                                string financialSoundness = "";
                                string globalCompetitiveness = "";
                                string innovation = "";
                                string longTermInvestmentValue = "";
                                string peopleManagement = "";
                                string qualityOfManagement = "";
                                string qualityOfProductsOrServices = "";
                                string socialResponsibility = "";
                                string useOfCorporateAssets = "";

                                HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                                htmlDoc.LoadHtml(fileText);
                                //HtmlNodeCollection baseInfoNodes = htmlDoc.DocumentNode.SelectNodes("//div[@class=\"company-info-card\"]/div[@class=\"row expanded\"]/div[@class=\"small-12 large-7 columns bg-white company-info-c\"]/div[@class=\"row company-info-card-table\"]/div[@class=\"row\"]");
                                HtmlNodeCollection infoNodes = htmlDoc.DocumentNode.SelectNodes("//table[@class=\"company-data-table\"]/tbody/tr");
                                foreach (HtmlNode baseInfoNode in infoNodes)
                                {
                                    HtmlNode infoTitleNode = baseInfoNode.SelectSingleNode("./th");
                                    HtmlNode infoValueNode = baseInfoNode.SelectSingleNode("./td");
                                    if (infoTitleNode != null && infoValueNode != null)
                                    {
                                        string infoTitle = CommonUtil.HtmlDecode(infoTitleNode.InnerText).Trim();
                                        string infoValue = CommonUtil.HtmlDecode(infoValueNode.InnerText).Trim();
                                        switch (infoTitle)
                                        {
                                            case "Industry":
                                                industry = infoValue;
                                                break;
                                            case "Industry Ranking":
                                                industryRank = infoValue;
                                                break;
                                            case "Overall Score":
                                                score = infoValue;
                                                break;
                                            case "Financial Soundness":
                                                financialSoundness = infoValue;
                                                break;
                                            case "Global Competitiveness":
                                                globalCompetitiveness = infoValue;
                                                break;
                                            case "Innovation":
                                                innovation = infoValue;
                                                break;
                                            case "Long-Term Investment Value":
                                                longTermInvestmentValue = infoValue;
                                                break;
                                            case "People Management":
                                                peopleManagement = infoValue;
                                                break;
                                            case "Quality of Management":
                                                qualityOfManagement = infoValue;
                                                break;
                                            case "Quality of Products/Services":
                                                qualityOfProductsOrServices = infoValue;
                                                break;
                                            case "Social Responsibility":
                                                socialResponsibility = infoValue;
                                                break;
                                            case "Use of Corporate Assets":
                                                useOfCorporateAssets = infoValue;
                                                break;
                                        }
                                    }
                                }

                                Dictionary<string, string> resultRow = new Dictionary<string, string>();
                                resultRow.Add("url", companyPageUrl);
                                resultRow.Add("companyName", listRow["companyName"]);
                                resultRow.Add("companyID", listRow["companyID"]);
                                resultRow.Add("year", listRow["yearValue"]);

                                resultRow.Add("industryName", industry);
                                resultRow.Add("industryRank", industryRank);
                                resultRow.Add("score", score);

                                resultRow.Add("financialSoundness", financialSoundness);
                                resultRow.Add("globalCompetitiveness", globalCompetitiveness);
                                resultRow.Add("innovation", innovation);
                                resultRow.Add("longTermInvestmentValue", longTermInvestmentValue);
                                resultRow.Add("peopleManagement", peopleManagement);
                                resultRow.Add("qualityOfManagement", qualityOfManagement);
                                resultRow.Add("qualityOfProductsOrServices", qualityOfProductsOrServices);
                                resultRow.Add("socialResponsibility", socialResponsibility);
                                resultRow.Add("useOfCorporateAssets", useOfCorporateAssets);

                                resultEW.AddRow(resultRow);
                            }
                            break;

                        case "2015":
                        case "2014":
                            {
                                Dictionary<string, string> resultRow = new Dictionary<string, string>();
                                resultRow.Add("url", companyPageUrl);
                                resultRow.Add("companyName", listRow["companyName"]);
                                resultRow.Add("companyID", listRow["companyID"]);
                                resultRow.Add("year", listRow["yearValue"]);

                                resultRow.Add("industryName", listRow["industry"]);
                                resultRow.Add("industryRank", listRow["industryRank"]);

                                resultRow.Add("financialSoundness", listRow["financialSoundness"]);
                                resultRow.Add("globalCompetitiveness", listRow["globalCompetitiveness"]);
                                resultRow.Add("innovation", listRow["innovation"]);
                                resultRow.Add("longTermInvestmentValue", listRow["longTermInvestmentValue"]);
                                resultRow.Add("peopleManagement", listRow["peopleManagement"]);
                                resultRow.Add("qualityOfManagement", listRow["qualityOfManagement"]);
                                resultRow.Add("qualityOfProductsOrServices", listRow["qualityOfProductsOrServices"]);
                                resultRow.Add("socialResponsibility", listRow["socialResponsibility"]);
                                resultRow.Add("useOfCorporateAssets", listRow["useOfCorporateAssets"]);

                                resultEW.AddRow(resultRow); 
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

            resultEW.SaveToDisk();
        }         

        private ExcelWriter CreateCompanyPageWriter()
        {
            String exportDir = this.RunPage.GetExportDir();
            string resultFilePath = Path.Combine(exportDir, "Fortune_FortuneCom_CompanyList.xlsx");

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
            resultColumnDic.Add("longTermInvestmentValue", 13);
            resultColumnDic.Add("qualityOfProductsOrServices", 14);
            resultColumnDic.Add("globalCompetitiveness", 15);


            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }         
    }
}