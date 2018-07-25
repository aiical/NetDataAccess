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
    public class GetListPages : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            this.GenerateListPageUrls(listSheet);
            return true;
        }

        private ExcelWriter GetExcelWriter(string destFilePath)
        { 

            Dictionary<string, int> columnDic = CommonUtil.InitStringIndexDic(new string[]{
                    "detailPageUrl",
                    "detailPageName", 
                    "cookie",
                    "grabStatus", 
                    "giveUpGrab", 
                    "Company_Name", 
                    "Page_Company_Name",
                    "Reviews", 
                    "Salaries", 
                    "InterViews"});

            ExcelWriter ew = new ExcelWriter(destFilePath, "List", columnDic);
            return ew;
        }

        private void GenerateListPageUrls(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            string resultFilePath = Path.Combine(exportDir, "GlassDoor_获取公司页.xlsx");

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;

            Dictionary<string, string> houseDic = new Dictionary<string, string>();
             
            ExcelWriter resultEW = this.GetExcelWriter(resultFilePath);

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
                            string reviews = infoJo.GetValue("reviewCount").ToString();
                            string salaries = infoJo.GetValue("salaryCount").ToString();
                            string interviews = infoJo.GetValue("interviewCount").ToString();

                            Dictionary<string, string> resultRow = new Dictionary<string, string>();
                            resultRow.Add("detailPageUrl", detailPageUrl);
                            resultRow.Add("detailPageName", companyName);
                            resultRow.Add("cookie", cookie);
                            resultRow.Add("Company_Name", companyName);
                            resultRow.Add("Page_Company_Name", pageCompanyName);
                            resultRow.Add("Reviews", reviews);
                            resultRow.Add("Salaries", salaries);
                            resultRow.Add("InterViews", interviews);
                            resultEW.AddRow(resultRow);
                        }
                        else
                        {
                            HtmlNodeCollection companyNodes = pageHtmlDoc.DocumentNode.SelectNodes("//div[@class=\"eiHdrModule module snug \"]");
                            if (companyNodes == null || companyNodes.Count == 0)
                            {
                                //获取公司列表页失败
                                //throw new Exception("获取公司列表页失败, url = " + url);
                                this.RunPage.InvokeAppendLogText("(" + (i + 1).ToString() + "/" + listSheet.RowCount.ToString() + ")获取公司列表页失败, url = " + url, LogLevelType.System, true);

                                /*
                                Dictionary<string, string> resultRow = new Dictionary<string, string>(); 
                                resultRow.Add("Company_Name", companyName); 
                                resultEW.AddRow(resultRow);
                                */
                            }
                            else
                            {
                                string linkUrl = "";
                                string page_Company_Name = "";
                                int reviewCount = 0;
                                int salaryCount = 0;
                                int interviewCount = 0;
                                string[] companyNameParts = companyName.ToLower().Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);

                                for (int j = 0; j < companyNodes.Count; j++)
                                {
                                    HtmlNode companyNode = companyNodes[j];
                                    HtmlNode linkNode = companyNode.SelectSingleNode("./div/div/div/a[@class=\"tightAll h2\"]");
                                    string tempCompanyName = CommonUtil.HtmlDecode(linkNode.InnerText).Trim();
                                    string lowCompanyName = tempCompanyName.ToLower();
                                    if (lowCompanyName.Contains(companyNameParts[0]))
                                    {
                                        HtmlNode reviewCountNode = companyNode.SelectSingleNode("./div/a[@class=\"eiCell cell reviews\"]/span[@class=\"num h2\"]");
                                        if (reviewCountNode != null)
                                        {
                                            string reviewCountText = CommonUtil.HtmlDecode(reviewCountNode.InnerText).Trim();
                                            int tempReviewCount = 0;
                                            if (reviewCountText == "--")
                                            {
                                                tempReviewCount = 0;
                                            }
                                            else if (reviewCountText.EndsWith("k"))
                                            {
                                                tempReviewCount = (int)(double.Parse(reviewCountText.Substring(0, reviewCountText.Length - 1)) * 1000);
                                            }
                                            else
                                            {
                                                tempReviewCount = int.Parse(reviewCountText);
                                            }
                                            if (tempReviewCount > reviewCount)
                                            {
                                                reviewCount = tempReviewCount;


                                                linkUrl = "https://www.glassdoor.com" + linkNode.GetAttributeValue("href", "");

                                                page_Company_Name = tempCompanyName;


                                                HtmlNode salaryCountNode = companyNode.SelectSingleNode("./div/a[@class=\"eiCell cell salaries\"]/span[@class=\"num h2\"]");
                                                if (salaryCountNode != null)
                                                {
                                                    string salaryCountText = CommonUtil.HtmlDecode(salaryCountNode.InnerText).Trim();
                                                    if (salaryCountText == "--")
                                                    {
                                                        salaryCount = 0;
                                                    }
                                                    else if (salaryCountText.EndsWith("k"))
                                                    {
                                                        salaryCount = (int)(double.Parse(salaryCountText.Substring(0, salaryCountText.Length - 1)) * 1000);
                                                    }
                                                    else
                                                    {
                                                        salaryCount = int.Parse(salaryCountText);
                                                    }
                                                }
                                                else
                                                {
                                                    interviewCount = 0;
                                                }


                                                HtmlNode interviewCountNode = companyNode.SelectSingleNode("./div/a[@class=\"eiCell cell interviews\"]/span[@class=\"num h2\"]");
                                                if (interviewCountNode != null)
                                                {
                                                    string interviewCountText = CommonUtil.HtmlDecode(interviewCountNode.InnerText).Trim();
                                                    if (interviewCountText == "--")
                                                    {
                                                        interviewCount = 0;
                                                    }
                                                    else if (interviewCountText.EndsWith("k"))
                                                    {
                                                        interviewCount = (int)(double.Parse(interviewCountText.Substring(0, interviewCountText.Length - 1)) * 1000);
                                                    }
                                                    else
                                                    {
                                                        interviewCount = int.Parse(interviewCountText);
                                                    }
                                                }
                                                else
                                                {
                                                    interviewCount = 0;
                                                }

                                            }
                                        }
                                    }
                                }

                                if (page_Company_Name.Length > 0)
                                {
                                    Dictionary<string, string> resultRow = new Dictionary<string, string>();
                                    resultRow.Add("detailPageUrl", linkUrl);
                                    resultRow.Add("detailPageName", companyName);
                                    resultRow.Add("cookie", cookie);
                                    resultRow.Add("Page_Company_Name", page_Company_Name);
                                    resultRow.Add("Company_Name", companyName);
                                    resultRow.Add("Reviews", reviewCount.ToString());
                                    resultRow.Add("Salaries", salaryCount.ToString());
                                    resultRow.Add("InterViews", interviewCount.ToString());
                                    resultEW.AddRow(resultRow);
                                }
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