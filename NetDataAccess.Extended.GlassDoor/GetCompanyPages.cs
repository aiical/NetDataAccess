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
    public class GetCompanyPages : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            this.GetCompanyInfos(listSheet);
            this.GetRatingPageUrls(listSheet);
            this.GetOverallDistributionPageUrls(listSheet);
            this.GetTrendPageUrls(listSheet);
            return true;
        }

        private ExcelWriter GetCompanInfoExcelWriter(string destFilePath)
        {

            Dictionary<string, int> columnDic = CommonUtil.InitStringIndexDic(new string[]{ 
                    "Company_Name", 
                    "Page_Company_Name",
                    "EmployerId",
                    "ReviewCount",
                    "JobCount", 
                    "SalaryCount", 
                    "InterViewCount", 
                    "BenefitCount", 
                    "PhotoCount", 
                    "WebSite", 
                    "Headquarters", 
                    "Size", 
                    "Founded", 
                    "Type", 
                    "Industry", 
                    "Revenue", 
                    "Competitors", 
                    "RatingNum", 
                    "RecommenToAFriend", 
                    "ApproveOfCEO", 
                    "CEOName", 
                    "CEORatings"});

            Dictionary<string, string> columnFormats = new Dictionary<string, string>();

            columnFormats.Add("ReviewCount", "#0");
            columnFormats.Add("JobCount", "#0");
            columnFormats.Add("SalaryCount", "#0");
            columnFormats.Add("InterViewCount", "#0");
            columnFormats.Add("BenefitCount", "#0");
            columnFormats.Add("PhotoCount", "#0");
            columnFormats.Add("RatingNum", "#0.00");
            columnFormats.Add("RecommenToAFriend", "#0.00");
            columnFormats.Add("ApproveOfCEO", "#0.00");
            columnFormats.Add("CEORatings", "#0");

            ExcelWriter ew = new ExcelWriter(destFilePath, "List", columnDic, columnFormats);
            return ew;
        }

        private void GetCompanyInfos(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            string resultFilePath = Path.Combine(exportDir, "GlassDoor_公司信息.xlsx");

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
                    string cookie = row["cookie"];

                    string pageFilePath = this.RunPage.GetFilePath(url, pageSourceDir);
                    string html = FileHelper.GetTextFromFile(pageFilePath);

                    HtmlAgilityPack.HtmlDocument pageHtmlDoc = new HtmlAgilityPack.HtmlDocument();
                    pageHtmlDoc.LoadHtml(html);

                    try
                    {
                        HtmlNode eiNode = pageHtmlDoc.DocumentNode.SelectSingleNode("//div[@id=\"EI\"]");
                        if (eiNode != null)
                        {
                            //获取列表页时直接获取了详情页
                            HtmlNode linkNode = pageHtmlDoc.DocumentNode.SelectSingleNode("//a[@class=\"sqLogoLink\"]");
                            HtmlNode titleNode = pageHtmlDoc.DocumentNode.SelectSingleNode("//h1[@class=\" strong tightAll\"]");
                            string detailPageUrl = "https://www.glassdoor.com" + linkNode.GetAttributeValue("href", "");
                            string pageCompanyName = CommonUtil.HtmlDecode(titleNode.GetAttributeValue("data-company", ""));
                            int beginIndex = html.IndexOf("sectionCounts");

                            int jsonBeginIndex = html.IndexOf("{", beginIndex);
                            int jsonEndIndex = html.IndexOf("}", beginIndex);
                            string jsonText = html.Substring(jsonBeginIndex, jsonEndIndex - jsonBeginIndex + 1);

                            JObject infoJo = JObject.Parse(jsonText);

                            string employerId = infoJo.GetValue("employerId").ToString();
                            string reviewCountStr = infoJo.GetValue("reviewCount").ToString();
                            string jobCountStr = infoJo.GetValue("jobCount").ToString();
                            string salaryCountStr = infoJo.GetValue("salaryCount").ToString();
                            string interviewCountStr = infoJo.GetValue("interviewCount").ToString();
                            string benefitCountStr = infoJo.GetValue("benefitCount").ToString();
                            string photoCountStr = infoJo.GetValue("photoCount").ToString();

                            string webSite = "";
                            string headquarters = "";
                            string size = "";
                            string founded = "";
                            string type = "";
                            string industry = "";
                            string revenue = "";
                            string competitors = "";

                            string ratingNumStr = "";
                            string recommenToAFriendStr = "";
                            string approveOfCEOStr = "";
                            string ceoName = "";
                            string ceoRatingsStr = "";

                            HtmlNodeCollection basicInfoNodes = pageHtmlDoc.DocumentNode.SelectNodes("//div[@id=\"EmpBasicInfo\"]/div/div/div[@class=\"infoEntity\"]");

                            if (basicInfoNodes != null)
                            {
                                foreach (HtmlNode basicInfoNode in basicInfoNodes)
                                {
                                    string label = CommonUtil.HtmlDecode(basicInfoNode.SelectSingleNode("./label").InnerText).Trim();
                                    string value = CommonUtil.HtmlDecode(basicInfoNode.SelectSingleNode("./span").InnerText).Trim();
                                    switch (label)
                                    {
                                        case "Website":
                                            webSite = value;
                                            break;
                                        case "Headquarters":
                                            headquarters = value;
                                            break;
                                        case "Size":
                                            size = value;
                                            break;
                                        case "Founded":
                                            founded = value;
                                            break;
                                        case "Type":
                                            type = value;
                                            break;
                                        case "Industry":
                                            industry = value;
                                            break;
                                        case "Revenue":
                                            revenue = value;
                                            break;
                                        case "Competitors":
                                            competitors = value;
                                            break;
                                    }
                                }
                            }

                            HtmlNode ratingNumNode = pageHtmlDoc.DocumentNode.SelectSingleNode("//div[@class=\"ratingNum\"]");
                            if (ratingNumNode != null)
                            {
                                ratingNumStr = CommonUtil.HtmlDecode(ratingNumNode.InnerText).Trim();
                            }


                            HtmlNode recommenToAFriendNode = pageHtmlDoc.DocumentNode.SelectSingleNode("//div[@id=\"EmpStats_Recommend\"]");
                            if (recommenToAFriendNode != null)
                            {
                                recommenToAFriendStr = CommonUtil.HtmlDecode(recommenToAFriendNode.GetAttributeValue("data-percentage", "")).Trim();
                            }

                            HtmlNode approveOfCEONode = pageHtmlDoc.DocumentNode.SelectSingleNode("//div[@id=\"EmpStats_Approve\"]");
                            if (approveOfCEONode != null)
                            {
                                approveOfCEOStr = CommonUtil.HtmlDecode(approveOfCEONode.GetAttributeValue("data-percentage", "")).Trim();
                            }

                            HtmlNodeCollection ceoNodes = pageHtmlDoc.DocumentNode.SelectNodes("//div[@class=\"empStatsBody\"]/div[@class=\"tbl gfxContainer\"]/div[last()]/div/div[@class=\"cell middle text\"]/div");
                            if (ceoNodes != null && ceoNodes.Count >= 2)
                            {
                                ceoName = CommonUtil.HtmlDecode(ceoNodes[0].InnerText).Trim();
                                ceoRatingsStr = CommonUtil.HtmlDecode(ceoNodes[1].InnerText).Trim().Replace(" Ratings", "").Replace(" Rating", "");
                            }
                            else
                            {
                                ceoNodes = pageHtmlDoc.DocumentNode.SelectNodes("//div[@class=\"empStatsBody\"]/div[@class=\"tbl gfxContainer\"]/div[last()]/a/div/div[@class=\"cell middle text\"]/div");
                                if (ceoNodes != null && ceoNodes.Count >= 2)
                                {
                                    ceoName = CommonUtil.HtmlDecode(ceoNodes[0].InnerText).Trim();
                                    ceoRatingsStr = CommonUtil.HtmlDecode(ceoNodes[1].InnerText).Trim().Replace(" Ratings", "").Replace(" Rating", "");
                                }
                            }

                            Dictionary<string, object> resultRow = new Dictionary<string, object>();
                            resultRow.Add("Company_Name", companyName);
                            resultRow.Add("Page_Company_Name", pageCompanyName);
                            resultRow.Add("EmployerId", employerId);
                            resultRow.Add("ReviewCount", reviewCountStr.Length == 0 ? null : (object)decimal.Parse(reviewCountStr));
                            resultRow.Add("JobCount", jobCountStr.Length == 0 ? null : (object)decimal.Parse(jobCountStr));
                            resultRow.Add("SalaryCount", salaryCountStr.Length == 0 ? null : (object)decimal.Parse(salaryCountStr));
                            resultRow.Add("InterViewCount", interviewCountStr.Length == 0 ? null : (object)decimal.Parse(interviewCountStr));
                            resultRow.Add("BenefitCount", benefitCountStr.Length == 0 ? null : (object)decimal.Parse(benefitCountStr));
                            resultRow.Add("PhotoCount", photoCountStr.Length == 0 ? null : (object)decimal.Parse(photoCountStr));
                            resultRow.Add("WebSite", webSite);
                            resultRow.Add("Headquarters", headquarters);
                            resultRow.Add("Size", size);
                            resultRow.Add("Founded", founded);
                            resultRow.Add("Type", type);
                            resultRow.Add("Industry", industry);
                            resultRow.Add("Revenue", revenue);
                            resultRow.Add("Competitors", competitors);
                            resultRow.Add("RatingNum", ratingNumStr.Length == 0 ? null : (object)decimal.Parse(ratingNumStr));
                            resultRow.Add("RecommenToAFriend", recommenToAFriendStr.Length == 0 ? null : (object)(decimal.Parse(recommenToAFriendStr) * (decimal)0.01));
                            resultRow.Add("ApproveOfCEO", approveOfCEOStr.Length == 0 ? null : (object)(decimal.Parse(approveOfCEOStr) * (decimal)0.01));
                            resultRow.Add("CEOName", ceoName);
                            resultRow.Add("CEORatings", ceoRatingsStr.Length == 0 ? null : (object)decimal.Parse(ceoRatingsStr));
                            resultEW.AddRow(resultRow);
                        }
                        else
                        {
                            throw new Exception("无法找到详情节点");
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

        private ExcelWriter GetRatingPageUrlsExcelWriter(string destFilePath)
        {

            Dictionary<string, int> columnDic = CommonUtil.InitStringIndexDic(new string[]{ 
                    "detailPageUrl",
                    "detailPageName", 
                    "cookie",
                    "grabStatus", 
                    "giveUpGrab", 
                    "Company_Name", 
                    "Page_Company_Name",
                    "EmployerId"});

            ExcelWriter ew = new ExcelWriter(destFilePath, "List", columnDic);
            return ew;
        }

        private void GetRatingPageUrls(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            string resultFilePath = Path.Combine(exportDir, "GlassDoor_Rating信息.xlsx");

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;

            Dictionary<string, string> houseDic = new Dictionary<string, string>();

            ExcelWriter resultEW = this.GetRatingPageUrlsExcelWriter(resultFilePath);

            for (int i = 0; i < listSheet.RowCount; i++)
            {

                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string url = row[detailPageUrlColumnName];
                    string companyName = row["Company_Name"];
                    string cookie = row["cookie"];

                    string pageFilePath = this.RunPage.GetFilePath(url, pageSourceDir);
                    string html = FileHelper.GetTextFromFile(pageFilePath);

                    HtmlAgilityPack.HtmlDocument pageHtmlDoc = new HtmlAgilityPack.HtmlDocument();
                    pageHtmlDoc.LoadHtml(html);

                    try
                    {
                        HtmlNode eiNode = pageHtmlDoc.DocumentNode.SelectSingleNode("//div[@id=\"EI\"]");
                        if (eiNode != null)
                        {
                            //获取列表页时直接获取了详情页
                            HtmlNode linkNode = pageHtmlDoc.DocumentNode.SelectSingleNode("//a[@class=\"sqLogoLink\"]");
                            HtmlNode titleNode = pageHtmlDoc.DocumentNode.SelectSingleNode("//h1[@class=\" strong tightAll\"]");
                            string detailPageUrl = "https://www.glassdoor.com" + linkNode.GetAttributeValue("href", "");
                            string pageCompanyName = CommonUtil.HtmlDecode(titleNode.GetAttributeValue("data-company", ""));
                            int beginIndex = html.IndexOf("sectionCounts");

                            int jsonBeginIndex = html.IndexOf("{", beginIndex);
                            int jsonEndIndex = html.IndexOf("}", beginIndex);
                            string jsonText = html.Substring(jsonBeginIndex, jsonEndIndex - jsonBeginIndex + 1);

                            JObject infoJo = JObject.Parse(jsonText);

                            string employerId = infoJo.GetValue("employerId").ToString();
                            string ratingPageUrl = "https://www.glassdoor.com/api/employer/" + employerId + "-rating.htm?locationStr=&jobTitleStr=&filterCurrentEmployee=false";

                            Dictionary<string, string> resultRow = new Dictionary<string, string>();
                            resultRow.Add("detailPageUrl", ratingPageUrl);
                            resultRow.Add("detailPageName", companyName);
                            resultRow.Add("cookie", cookie);
                            resultRow.Add("Company_Name", companyName);
                            resultRow.Add("Page_Company_Name", pageCompanyName);
                            resultRow.Add("EmployerId", employerId);
                            resultEW.AddRow(resultRow);
                        }
                        else
                        {
                            throw new Exception("无法找到详情节点");
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

        private ExcelWriter GetOverallDistributioPageUrlsExcelWriter(string destFilePath)
        {

            Dictionary<string, int> columnDic = CommonUtil.InitStringIndexDic(new string[]{ 
                    "detailPageUrl",
                    "detailPageName", 
                    "cookie",
                    "grabStatus", 
                    "giveUpGrab", 
                    "Company_Name", 
                    "Page_Company_Name",
                    "EmployerId"});

            ExcelWriter ew = new ExcelWriter(destFilePath, "List", columnDic);
            return ew;
        }

        private void GetOverallDistributionPageUrls(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            string resultFilePath = Path.Combine(exportDir, "GlassDoor_OverallDistribution信息.xlsx");

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;

            Dictionary<string, string> houseDic = new Dictionary<string, string>();

            ExcelWriter resultEW = this.GetOverallDistributioPageUrlsExcelWriter(resultFilePath);

            for (int i = 0; i < listSheet.RowCount; i++)
            {

                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string url = row[detailPageUrlColumnName];
                    string companyName = row["Company_Name"];
                    string cookie = row["cookie"];

                    string pageFilePath = this.RunPage.GetFilePath(url, pageSourceDir);
                    string html = FileHelper.GetTextFromFile(pageFilePath);

                    HtmlAgilityPack.HtmlDocument pageHtmlDoc = new HtmlAgilityPack.HtmlDocument();
                    pageHtmlDoc.LoadHtml(html);

                    try
                    {
                        HtmlNode eiNode = pageHtmlDoc.DocumentNode.SelectSingleNode("//div[@id=\"EI\"]");
                        if (eiNode != null)
                        {
                            //获取列表页时直接获取了详情页
                            HtmlNode linkNode = pageHtmlDoc.DocumentNode.SelectSingleNode("//a[@class=\"sqLogoLink\"]");
                            HtmlNode titleNode = pageHtmlDoc.DocumentNode.SelectSingleNode("//h1[@class=\" strong tightAll\"]");
                            string detailPageUrl = "https://www.glassdoor.com" + linkNode.GetAttributeValue("href", "");
                            string pageCompanyName = CommonUtil.HtmlDecode(titleNode.GetAttributeValue("data-company", ""));
                            int beginIndex = html.IndexOf("sectionCounts");

                            int jsonBeginIndex = html.IndexOf("{", beginIndex);
                            int jsonEndIndex = html.IndexOf("}", beginIndex);
                            string jsonText = html.Substring(jsonBeginIndex, jsonEndIndex - jsonBeginIndex + 1);

                            JObject infoJo = JObject.Parse(jsonText);

                            string employerId = infoJo.GetValue("employerId").ToString();
                            string odPageUrl = "https://www.glassdoor.com/api/employer/" + employerId + "-rating.htm?dataType=distribution&category=overallRating";

                            Dictionary<string, string> resultRow = new Dictionary<string, string>();
                            resultRow.Add("detailPageUrl", odPageUrl);
                            resultRow.Add("detailPageName", companyName);
                            resultRow.Add("cookie", cookie);
                            resultRow.Add("Company_Name", companyName);
                            resultRow.Add("Page_Company_Name", pageCompanyName);
                            resultRow.Add("EmployerId", employerId);
                            resultEW.AddRow(resultRow);
                        }
                        else
                        {
                            throw new Exception("无法找到详情节点");
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

        private ExcelWriter GetTrendPageUrlsExcelWriter(string destFilePath)
        {

            Dictionary<string, int> columnDic = CommonUtil.InitStringIndexDic(new string[]{ 
                    "detailPageUrl",
                    "detailPageName", 
                    "cookie",
                    "grabStatus", 
                    "giveUpGrab", 
                    "Company_Name", 
                    "Page_Company_Name",
                    "EmployerId"});

            ExcelWriter ew = new ExcelWriter(destFilePath, "List", columnDic);
            return ew;
        }

        private void GetTrendPageUrls(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            string resultFilePath = Path.Combine(exportDir, "GlassDoor_Trend信息.xlsx");

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;

            Dictionary<string, string> houseDic = new Dictionary<string, string>();

            ExcelWriter resultEW = this.GetTrendPageUrlsExcelWriter(resultFilePath);

            for (int i = 0; i < listSheet.RowCount; i++)
            {

                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string url = row[detailPageUrlColumnName];
                    string companyName = row["Company_Name"];
                    string cookie = row["cookie"];

                    string pageFilePath = this.RunPage.GetFilePath(url, pageSourceDir);
                    string html = FileHelper.GetTextFromFile(pageFilePath);

                    HtmlAgilityPack.HtmlDocument pageHtmlDoc = new HtmlAgilityPack.HtmlDocument();
                    pageHtmlDoc.LoadHtml(html);

                    try
                    {
                        HtmlNode eiNode = pageHtmlDoc.DocumentNode.SelectSingleNode("//div[@id=\"EI\"]");
                        if (eiNode != null)
                        {
                            //获取列表页时直接获取了详情页
                            HtmlNode linkNode = pageHtmlDoc.DocumentNode.SelectSingleNode("//a[@class=\"sqLogoLink\"]");
                            HtmlNode titleNode = pageHtmlDoc.DocumentNode.SelectSingleNode("//h1[@class=\" strong tightAll\"]");
                            string detailPageUrl = "https://www.glassdoor.com" + linkNode.GetAttributeValue("href", "");
                            string pageCompanyName = CommonUtil.HtmlDecode(titleNode.GetAttributeValue("data-company", ""));
                            int beginIndex = html.IndexOf("sectionCounts");

                            int jsonBeginIndex = html.IndexOf("{", beginIndex);
                            int jsonEndIndex = html.IndexOf("}", beginIndex);
                            string jsonText = html.Substring(jsonBeginIndex, jsonEndIndex - jsonBeginIndex + 1);

                            JObject infoJo = JObject.Parse(jsonText);

                            string employerId = infoJo.GetValue("employerId").ToString();
                            string odPageUrl = "https://www.glassdoor.com/api/employer/" + employerId + "-rating.htm?dataType=trend&category=overallRating&locationStr=&jobTitleStr=&filterCurrentEmployee=false";

                            Dictionary<string, string> resultRow = new Dictionary<string, string>();
                            resultRow.Add("detailPageUrl", odPageUrl);
                            resultRow.Add("detailPageName", companyName);
                            resultRow.Add("cookie", cookie);
                            resultRow.Add("Company_Name", companyName);
                            resultRow.Add("Page_Company_Name", pageCompanyName);
                            resultRow.Add("EmployerId", employerId);
                            resultEW.AddRow(resultRow);
                        }
                        else
                        {
                            throw new Exception("无法找到详情节点");
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