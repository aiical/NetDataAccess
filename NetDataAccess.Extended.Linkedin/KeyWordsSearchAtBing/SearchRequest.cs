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
using System.Web;
using System.Runtime.Remoting;
using System.Reflection;
using System.Collections;

namespace NetDataAccess.Extended.Linkedin.KeyWordsSearchAtBing
{ 
    public class SearchRequest : ExternalRunWebPage
    {
        public override void CheckRequestCompleteFile(string webPageText, Dictionary<string, string> listRow)
        {
            HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
            htmlDoc.LoadHtml(webPageText);

            if (!webPageText.Contains("There are no results for"))
            {
                HtmlNodeCollection liElements = htmlDoc.DocumentNode.SelectNodes("//li[@class=\"b_algo\"]");
                if (liElements == null || liElements.Count == 0)
                {
                    throw new Exception("返回的文本页面不准确.");
                }
            }
        }

        public override bool AfterAllGrab(IListSheet listSheet)
        {
            try
            {

                this.GetLinkedinUrls(listSheet);
                return base.AfterAllGrab(listSheet);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void GetLinkedinUrls(IListSheet listSheet)
        {
            ExcelWriter matchedUrlEW = this.GetExcelWriter();
            ExcelWriter unmatchedNameEW = this.GetNoneMatchedExcelWriter();

            int rowCount = listSheet.GetListDBRowCount();
            for (int i = 0; i < rowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);

                HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);

                HtmlNodeCollection liElements = htmlDoc.DocumentNode.SelectNodes("//li[@class=\"b_algo\"]");
                if (liElements == null || liElements.Count == 0)
                {
                    Dictionary<string, string> resultRow = new Dictionary<string, string>();
                    resultRow.Add("FirmID", row["FirmID"]);
                    resultRow.Add("FirmName", row["FirmName"]);
                    resultRow.Add("LastName", row["LastName"]);
                    resultRow.Add("FirstName", row["FirstName"]);
                    resultRow.Add("MiddleName", row["MiddleName"]);
                    resultRow.Add("MatchType", "NoResults");

                    unmatchedNameEW.AddRow(resultRow);
                }
                else
                {
                    int sameNameCount = 0;
                    for (int j = 0; j < liElements.Count; j++)
                    {
                        HtmlNode linkNode = liElements[j].SelectSingleNode("./div[@class=\"b_title\"]/h2/a");
                        if (linkNode == null)
                        {
                            linkNode = liElements[j].SelectSingleNode("./h2/a");
                        }
                        string linkUrl = linkNode.GetAttributeValue("href", "");
                        if (linkUrl.EndsWith("/zh-cn"))
                        {
                            linkUrl = linkUrl.Substring(0, linkUrl.Length - "/zh-cn".Length);
                        }
                        string linkText = CommonUtil.HtmlDecode(linkNode.InnerText.Trim());
                        //重名的
                        if (this.CheckInText(linkText, row["FirstName"]) && this.CheckInText(linkText, row["LastName"]))
                        {
                            sameNameCount++;

                            Dictionary<string, string> resultUrlRow = new Dictionary<string, string>();
                            string detailPageUrl = linkUrl + "?tt=" + i.ToString() + "_" + j.ToString();
                            resultUrlRow.Add("detailPageUrl", linkUrl);
                            resultUrlRow.Add("detailPageName", detailPageUrl);
                            resultUrlRow.Add("FirmID", row["FirmID"]);
                            resultUrlRow.Add("FirmName", row["FirmName"]);
                            resultUrlRow.Add("LastName", row["LastName"]);
                            resultUrlRow.Add("FirstName", row["FirstName"]);
                            resultUrlRow.Add("MiddleName", row["MiddleName"]);

                            matchedUrlEW.AddRow(resultUrlRow);
                        }
                    }

                    if (sameNameCount == 0)
                    {
                        for (int j = 0; j < liElements.Count; j++)
                        {
                            HtmlNode linkNode = liElements[j].SelectSingleNode("./div[@class=\"b_title\"]/h2/a");
                            if (linkNode == null)
                            {
                                linkNode = liElements[j].SelectSingleNode("./h2/a");
                            }
                            string linkUrl = linkNode.GetAttributeValue("href", "");
                            if (linkUrl.EndsWith("/zh-cn"))
                            {
                                linkUrl = linkUrl.Substring(0, linkUrl.Length - "/zh-cn".Length);
                            }
                            string linkText = CommonUtil.HtmlDecode(linkNode.InnerText.Trim());
                            //姓相同，名的第一个字母相同
                            if (this.CheckInText(linkText, row["FirstName"].Substring(0, 1)) && this.CheckInText(linkText, row["LastName"]))
                            {
                                sameNameCount++;

                                Dictionary<string, string> resultUrlRow = new Dictionary<string, string>();
                                string detailPageUrl = linkUrl + "?tt=" + i.ToString() + "_" + j.ToString();
                                resultUrlRow.Add("detailPageUrl", linkUrl);
                                resultUrlRow.Add("detailPageName", detailPageUrl);
                                resultUrlRow.Add("FirmID", row["FirmID"]);
                                resultUrlRow.Add("FirmName", row["FirmName"]);
                                resultUrlRow.Add("LastName", row["LastName"]);
                                resultUrlRow.Add("FirstName", row["FirstName"]);
                                resultUrlRow.Add("MiddleName", row["MiddleName"]);

                                matchedUrlEW.AddRow(resultUrlRow);
                            }
                        }
                    }


                    Dictionary<string, string> resultRow = new Dictionary<string, string>();
                    resultRow.Add("FirmID", row["FirmID"]);
                    resultRow.Add("FirmName", row["FirmName"]);
                    resultRow.Add("LastName", row["LastName"]);
                    resultRow.Add("FirstName", row["FirstName"]);
                    resultRow.Add("MiddleName", row["MiddleName"]);
                    resultRow.Add("MatchType", sameNameCount == 0 ? "NoSameNameResults" : "HasResults");

                    unmatchedNameEW.AddRow(resultRow);
                }
            }
            matchedUrlEW.SaveToDisk();
            unmatchedNameEW.SaveToDisk();
        }

        private bool CheckInText(string sourceText, string partText)
        {
            string sourceTextLower = sourceText.ToLower();
            string partTextLower = partText.ToLower();
            string[] parts = partTextLower.Split(new string[] {" " }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < parts.Length; i++)
            {
                if (!sourceTextLower.Contains(parts[i]))
                {
                    return false;
                }
            }
            return true;
        }

        private ExcelWriter GetExcelWriter()
        {
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                "detailPageUrl",
                "detailPageName", 
                "cookie",
                "grabStatus", 
                "giveUpGrab", 
                "FirmID",
                "FirmName",
                "LastName",
                "FirstName",
                "MiddleName"});
            string resultFilePath = Path.Combine(exportDir, "Linkedin个人详情页.xlsx");

            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic);
            return resultEW;
        }


        private ExcelWriter GetNoneMatchedExcelWriter()
        {
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                "FirmID",
                "FirmName",
                "LastName",
                "FirstName",
                "MiddleName",
                "MatchType"});
            string resultFilePath = Path.Combine(exportDir, "Linkedin自动匹配结果.xlsx");

            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic);
            return resultEW;
        }
    }
}