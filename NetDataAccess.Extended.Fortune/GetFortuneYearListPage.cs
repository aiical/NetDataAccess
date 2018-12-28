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
    public class GetFortuneYearListPage : ExternalRunWebPage
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
                            case "2018":
                                {
                                    JObject rootObj = JObject.Parse(fileText);
                                    JArray companyJArray = rootObj.GetValue("list-items") as JArray;
                                    for (int j = 0; j < companyJArray.Count; j++)
                                    {
                                        JObject companyObj = companyJArray[j] as JObject;
                                        string companyName = companyObj.GetValue("title").ToString();
                                        string companyID = companyObj.GetValue("id").ToString();
                                        string companyPageUrl = companyObj.GetValue("permalink").ToString().Replace("*", "");
                                        int endIndex = companyPageUrl.LastIndexOf("-");
                                        companyPageUrl = companyPageUrl.Substring(0, endIndex);

                                        Dictionary<string, string> resultRow = new Dictionary<string, string>();
                                        resultRow.Add("detailPageUrl", companyPageUrl);
                                        resultRow.Add("detailPageName", companyPageUrl);
                                        resultRow.Add("companyName", companyName);
                                        resultRow.Add("companyID", companyID);
                                        resultRow.Add("yearValue", year);
                                        resultEW.AddRow(resultRow);
                                    }
                                }
                                break;
                            case "2017": 
                                {
                                    JObject rootObj = JObject.Parse(fileText);
                                    JArray companyJArray = rootObj.GetValue("list-items") as JArray;
                                    for (int j = 0; j < companyJArray.Count; j++)
                                    {
                                        JObject companyObj = companyJArray[j] as JObject;
                                        string companyName = companyObj.GetValue("title").ToString();
                                        string companyID = companyObj.GetValue("id").ToString();
                                        string companyPageUrl = companyObj.GetValue("permalink").ToString().Replace("*", "");

                                        Dictionary<string, string> resultRow = new Dictionary<string, string>();
                                        resultRow.Add("detailPageUrl", companyPageUrl);
                                        resultRow.Add("detailPageName", companyPageUrl);
                                        resultRow.Add("companyName", companyName);
                                        resultRow.Add("companyID", companyID);
                                        resultRow.Add("yearValue", year);
                                        resultEW.AddRow(resultRow);
                                    }
                                }
                                break; 
                            case "2016":
                                {
                                    string beginStr = "var fortune_wp_vars = ";
                                    string endStr = "}}}}};";
                                    int beginIndex = fileText.IndexOf(beginStr);
                                    int endIndex = fileText.IndexOf(endStr);
                                    string jsonText = fileText.Substring(beginIndex + beginStr.Length, endIndex + endStr.Length - beginIndex - beginStr.Length - 1);
                                    JObject rootObj = JObject.Parse(jsonText);
                                    JObject bootstrapObj = rootObj.GetValue("bootstrap") as JObject;
                                    JObject franchiseObj = bootstrapObj.GetValue("franchise") as JObject;
                                    JArray filtered_sorted_dataJArray = franchiseObj.GetValue("filtered_sorted_data") as JArray;
                                    for (int j = 0; j < filtered_sorted_dataJArray.Count; j++)
                                    {
                                        JObject companyObj = filtered_sorted_dataJArray[j] as JObject;
                                        string companyName = companyObj.GetValue("title").ToString();
                                        string companyID = companyObj.GetValue("id").ToString();
                                        string companyPageUrl = companyObj.GetValue("permalink").ToString().Replace("*", "");
                                        
                                        Dictionary<string, string> resultRow = new Dictionary<string, string>();
                                        resultRow.Add("detailPageUrl", companyPageUrl);
                                        resultRow.Add("detailPageName", companyPageUrl);
                                        resultRow.Add("companyName", companyName);
                                        resultRow.Add("companyID", companyID);
                                        resultRow.Add("yearValue", year);
                                        resultEW.AddRow(resultRow);
                                    }
                                }
                                break;

                            case "2015":
                            case "2014":
                                {
                                    JObject rootObj = JObject.Parse(fileText);
                                    JArray companyJArray = rootObj.GetValue("articles") as JArray;
                                    for (int j = 0; j < companyJArray.Count; j++)
                                    {
                                        JObject companyObj = companyJArray[j] as JObject;
                                        string companyName = companyObj.GetValue("title").ToString();
                                        string companyID = companyObj.GetValue("id").ToString();
                                        string companyPageUrl = companyObj.GetValue("url").ToString().Replace("*", "");
                                        int endIndex = companyPageUrl.LastIndexOf("-");
                                        companyPageUrl = companyPageUrl.Substring(0, endIndex);

                                        JObject highlightsObj = companyObj.GetValue("highlights") as JObject;
                                        string industry = highlightsObj.GetValue("Industry").ToString();
                                        string industryRank = highlightsObj.GetValue("Industry Rank").ToString();


                                        string financialSoundness = "";
                                        string globalCompetitiveness = "";
                                        string innovation = "";
                                        string longTermInvestmentValue = "";
                                        string peopleManagement = "";
                                        string qualityOfManagement = "";
                                        string qualityOfProductsOrServices = "";
                                        string socialResponsibility = "";
                                        string useOfCorporateAssets = "";
                                        JObject dataObj = ((companyObj.GetValue("tables") as JObject).GetValue("Nine Key Attributes of Reputation") as JObject).GetValue("data") as JObject;
                                        financialSoundness = (dataObj.GetValue("Financial soundness") as JArray)[0].ToString();
                                        globalCompetitiveness = (dataObj.GetValue("Global competitiveness") as JArray)[0].ToString();
                                        innovation = (dataObj.GetValue("Innovation") as JArray)[0].ToString();
                                        longTermInvestmentValue = (dataObj.GetValue("Long-term investment value") as JArray)[0].ToString();
                                        peopleManagement = (dataObj.GetValue("People management") as JArray)[0].ToString();
                                        qualityOfManagement = (dataObj.GetValue("Quality of management") as JArray)[0].ToString();
                                        qualityOfProductsOrServices = (dataObj.GetValue("Quality of products / services") as JArray)[0].ToString();
                                        socialResponsibility = (dataObj.GetValue("Social responsibility") as JArray)[0].ToString();
                                        useOfCorporateAssets = (dataObj.GetValue("Use of corporate assets") as JArray)[0].ToString();

                                        Dictionary<string, string> resultRow = new Dictionary<string, string>();
                                        resultRow.Add("detailPageUrl", companyPageUrl);
                                        resultRow.Add("detailPageName", companyPageUrl);
                                        resultRow.Add("giveUpGrab", "Y");
                                        resultRow.Add("companyName", companyName);
                                        resultRow.Add("companyID", companyID);
                                        resultRow.Add("yearValue", year);

                                        resultRow.Add("industry", industry);
                                        resultRow.Add("industryRank", industryRank);

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
            string resultFilePath = Path.Combine(exportDir, "Fortune_FortuneCom_CompanyDetail.xlsx");

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("yearValue", 5);
            resultColumnDic.Add("companyName", 6);
            resultColumnDic.Add("companyID", 7);
            resultColumnDic.Add("industry", 8);
            resultColumnDic.Add("industryRank", 9);

            resultColumnDic.Add("financialSoundness", 10);
            resultColumnDic.Add("globalCompetitiveness", 11);
            resultColumnDic.Add("innovation", 12);
            resultColumnDic.Add("longTermInvestmentValue", 13);
            resultColumnDic.Add("peopleManagement", 14);
            resultColumnDic.Add("qualityOfManagement", 15);
            resultColumnDic.Add("qualityOfProductsOrServices", 16);
            resultColumnDic.Add("socialResponsibility", 17);
            resultColumnDic.Add("useOfCorporateAssets", 18); 


            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }         
    }
}