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
using NetDataAccess.Base.CsvHelper;
using NetDataAccess.Base.Reader;

namespace NetDataAccess.Extended.GlassDoor
{
    public class GetReviewListPages : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            this.GetCompanyInfos(listSheet);
            return true;
        }

        private ExcelWriter GetCompanInfoExcelWriter(string destFilePath)
        {

            Dictionary<string, int> columnDic = CommonUtil.InitStringIndexDic(new string[]{ 
                    "Company_Name", 
                    "Page_Company_Name", 
                    "ReviewId",
                    "PostDate",
                    "PostTime",
                    "Summary",
                    "Rating", 
                    "WorkLifeBalance",
                    "CultureValues",
                    "CareerOpportunities",
                    "CompBenefits",
                    "SeniorManagement",
                    "Employee", 
                    "Job",
                    "Location", 
                    "Recommends", 
                    "Outlook", 
                    "OptionOfCEO", 
                    "WorkingTime", 
                    "Pros", 
                    "Cons", 
                    "AdviceToManagement"});

            Dictionary<string, string> columnFormats = new Dictionary<string, string>();

            columnFormats.Add("Rating", "#0.00");

            ExcelWriter ew = new ExcelWriter(destFilePath, "List", columnDic, columnFormats);
            return ew;
        }

        private void GetCompanyInfos(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            string resultFilePath = Path.Combine(exportDir, "GlassDoor_ReviewDetail.xlsx");

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;

            Dictionary<string, string> houseDic = new Dictionary<string, string>();

            ExcelWriter resultEW = this.GetCompanInfoExcelWriter(resultFilePath);

            for (int i = 0; i < listSheet.RowCount; i++)
            {

                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string url = row[detailPageUrlColumnName];
                    string companyName = row["Company_Name"];
                    string pageCompanyName = row["Page_Company_Name"];

                    string pageFilePath = this.RunPage.GetFilePath(url, pageSourceDir);
                    string html = FileHelper.GetTextFromFile(pageFilePath);

                    HtmlAgilityPack.HtmlDocument pageHtmlDoc = new HtmlAgilityPack.HtmlDocument();
                    pageHtmlDoc.LoadHtml(html);

                    try
                    {
                        HtmlNodeCollection itemNodes = pageHtmlDoc.DocumentNode.SelectNodes("//li[@class=\" empReview cf \"]");
                        if (itemNodes != null)
                        {
                            foreach (HtmlNode itemNode in itemNodes)
                            {
                                string reviewId = CommonUtil.HtmlDecode( itemNode.GetAttributeValue("id","")).Trim();

                                HtmlNode dateNode = itemNode.SelectSingleNode("./div[@class=\"hreview\"]/div[@class=\"cf\"]/div[@class=\"floatLt\"]/time");
                                string postDate = dateNode == null ? "" : dateNode.GetAttributeValue("datetime", "");

                                HtmlNode summaryNode = itemNode.SelectSingleNode("./div[@class=\"hreview\"]/div[@class=\" tbl fill reviewTop\"]/div[@class=\"row\"]/div[@class=\"cell\"]/h2");
                                string summary = CommonUtil.HtmlDecode(summaryNode.InnerText).Trim();

                                HtmlNode reviewMetaNode = itemNode.SelectSingleNode("./div[@class=\"hreview\"]/div[@class=\" tbl fill reviewTop\"]/div[@class=\"row\"]/div[@class=\"cell\"]/div[@class=\"tbl reviewMeta\"]");

                                HtmlNode ratingNode = reviewMetaNode.SelectSingleNode("./div[@class=\"gdStarsWrapper cell top\"]/span[@class=\"gdStars gdRatings sm margRt\"]/span[@class=\"rating\"]/span");
                                decimal rating = decimal.Parse(ratingNode.GetAttributeValue("title", ""));

                                string workLifeBalance = ""; 
                                string cultureValues = "";
                                string careerOpportunities = "";
                                string compBenefits = "";
                                string seniorManagement = "";

                                HtmlNodeCollection subRatingLiNodes = reviewMetaNode.SelectNodes("./div[@class=\"gdStarsWrapper cell top\"]/span[@class=\"gdStars gdRatings sm margRt\"]/div[@class=\"subRatings module\"]/ul/li");
                                if (subRatingLiNodes != null)
                                {
                                    foreach (HtmlNode subRatingLiNode in subRatingLiNodes)
                                    {
                                        string key = CommonUtil.HtmlDecode(subRatingLiNode.SelectSingleNode("./div").InnerText).Trim();
                                        string value = CommonUtil.HtmlDecode(subRatingLiNode.SelectSingleNode("./span").GetAttributeValue("title", "")).Trim();

                                        switch (key)
                                        {
                                            case "Work/Life Balance":
                                                workLifeBalance = value;
                                                break;
                                            case "Culture & Values":
                                                cultureValues = value;
                                                break;
                                            case "Career Opportunities":
                                                careerOpportunities = value;
                                                break;
                                            case "Comp & Benefits":
                                                compBenefits = value;
                                                break;
                                            case "Senior Management":
                                                seniorManagement = value;
                                                break;
                                        }
                                    }
                                }

                                HtmlNode employeeNode = reviewMetaNode.SelectSingleNode("./div[@class=\" cell top\"]/div[@class=\"author minor tbl\"]/span[@class=\"authorInfo tbl hideHH\"]/span[@class=\"authorJobTitle middle reviewer\"]");
                                string employeeInfo = CommonUtil.HtmlDecode(employeeNode.InnerText).Trim();
                                string employee = "";
                                if (employeeInfo.Contains("Former Employee"))
                                {
                                    employee = "Former Employee";
                                }
                                else if (employeeInfo.Contains("Current Employee"))
                                {
                                    employee = "Current Employee";
                                }

                                int jobBeginIndex = employeeInfo.IndexOf("-");
                                string job = employeeInfo.Substring(jobBeginIndex + 1).Trim();



                                HtmlNode locationNode = reviewMetaNode.SelectSingleNode("./div[@class=\" cell top\"]/div[@class=\"author minor tbl\"]/span[@class=\"authorInfo tbl hideHH\"]/span[@class=\"authorLocation middle\"]");
                                string location = locationNode == null ? "" : CommonUtil.HtmlDecode(locationNode.InnerText).Trim();


                                HtmlNode detailNode = itemNode.SelectSingleNode("./div[@class=\"hreview\"]/div[@class=\"tbl fill\"]/div[@class=\"row\"]/div[@class=\"cell reviewBodyCell\"]");

                                HtmlNodeCollection tightNodes = detailNode.SelectNodes("./div[@class=\"flex-grid recommends\"]/div[@class=\"tightLt col span-1-3\"]");
                                
                                string recommends = "";
                                string outlook="";
                                string optionOfCEO = "";
                                if (tightNodes != null)
                                {
                                    foreach (HtmlNode tightNode in tightNodes)
                                    {
                                        string value = CommonUtil.HtmlDecode(tightNode.InnerText).Trim();
                                        if (value == "Recommends")
                                        {
                                            recommends = value;
                                        }
                                        else if (value.Contains("Outlook"))
                                        {
                                            outlook = value;
                                        }
                                        else if (value.Contains("CEO"))
                                        {
                                            optionOfCEO = value;
                                        }
                                    }
                                }

                                HtmlNode workingTimeNode = detailNode.SelectSingleNode("./p");
                                string workingTime =workingTimeNode ==null?"":  CommonUtil.HtmlDecode(workingTimeNode.InnerText).Trim();

                                string pros="";
                                string cons="";
                                string adviceToManagement="";
                                HtmlNodeCollection descriptionNodes = detailNode.SelectNodes("./div[@class=\"description \"]/div[@class=\" tbl fill prosConsAdvice truncateData\"]/div/div");
                                if (descriptionNodes != null)
                                {
                                    foreach (HtmlNode desNode in descriptionNodes)
                                    {
                                        HtmlNodeCollection pNodes = desNode.SelectNodes("./p");
                                        string desType = CommonUtil.HtmlDecode(pNodes[0].InnerText).Trim();
                                        string value = CommonUtil.HtmlDecode(pNodes[1].InnerText).Trim();
                                        switch (desType)
                                        {
                                            case "Pros":
                                                pros = value;
                                                break;
                                            case "Cons":
                                                cons = value;
                                                break;
                                            case "Advice to Management":
                                                adviceToManagement = value;
                                                break;
                                        }
                                    }
                                }

                                HtmlNode timeNode = detailNode.SelectSingleNode("./div[@class=\"description \"]/div[@class=\"tbl fill outlookEmpReview\"]/span[\"hidden\"]/span[@class=\"dtreviewed\"]");
                                string postTime = timeNode == null ? "" : CommonUtil.HtmlDecode(timeNode.InnerText).Trim();

                                Dictionary<string, object> resultRow = new Dictionary<string, object>();
                                resultRow.Add("Company_Name", companyName);
                                resultRow.Add("Page_Company_Name", pageCompanyName);
                                resultRow.Add("ReviewId", reviewId);
                                resultRow.Add("PostDate", postDate);
                                resultRow.Add("PostTime", postTime); 
                                resultRow.Add("Summary", summary); 
                                resultRow.Add("Rating", rating);
                                resultRow.Add("WorkLifeBalance", workLifeBalance);
                                resultRow.Add("CultureValues", cultureValues);
                                resultRow.Add("CareerOpportunities", careerOpportunities);
                                resultRow.Add("CompBenefits", compBenefits);
                                resultRow.Add("SeniorManagement", seniorManagement); 
                                resultRow.Add("Employee", employee); 
                                resultRow.Add("Job", job); 
                                resultRow.Add("Location", location); 
                                resultRow.Add("Recommends", recommends); 
                                resultRow.Add("Outlook", outlook); 
                                resultRow.Add("OptionOfCEO", optionOfCEO); 
                                resultRow.Add("WorkingTime", workingTime); 
                                resultRow.Add("Pros", pros); 
                                resultRow.Add("Cons", cons); 
                                resultRow.Add("AdviceToManagement", adviceToManagement); 
                                resultEW.AddRow(resultRow); 
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        this.RunPage.InvokeAppendLogText(ex.Message + ", pageUrl = " + url, LogLevelType.System, true);
                        throw ex;
                    }
                }
            }

            resultEW.SaveToDisk();
        } 
    }
}