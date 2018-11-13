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

namespace NetDataAccess.Extended.Keywords.WanFang
{
    public class GetKeywordSearchFirstPages : ExternalRunWebPage
    { 
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            try
            { 
                this.GetAllPages(listSheet);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return true;
        }

        public override void CheckRequestCompleteFile(string webPageText, Dictionary<string, string> listRow)
        {
            //如果是首页（即不包含“首页”俩字，那么才判断
            if (!webPageText.Contains("首 页"))
            {
                if (!webPageText.Contains("命中"))
                {
                    throw new Exception("下载不成功.");
                }
                else
                {
                    string nextPageUrl = this.GetNextPageUrlByHtml(webPageText);
                    if (nextPageUrl != null)
                    {
                        this.GetPageFile(nextPageUrl, listRow);
                    }
                }
            }
        }

        private void GetNextPage(string pageUrl, Dictionary<string, string> listRow)
        {
            string nextPageUrl = this.GetNextPageUrl(pageUrl);
            if (nextPageUrl != null)
            {
                this.GetPageFile(nextPageUrl, listRow);
            }
        }

        private string GetNextPageUrl(string pageUrl)
        {
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            string filePath = this.RunPage.GetFilePath(pageUrl, pageSourceDir);
            string webPageText = FileHelper.GetTextFromFile(filePath);
            return this.GetNextPageUrlByHtml(webPageText);
        }         

        private string GetNextPageUrlByHtml(string webPageText)
        { 
            HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
            htmlDoc.LoadHtml(webPageText);
            HtmlNode nextPageNode = htmlDoc.DocumentNode.SelectSingleNode("//a[@class=\"page\"]/t[contains(text(),\"下一页\")]");
            if (nextPageNode != null)
            {
                string nextPageUrl = "http://librarian.wanfangdata.com.cn/SearchResult.aspx" + CommonUtil.UrlDecodeSymbolAnd(nextPageNode.ParentNode.GetAttributeValue("href", ""));
                return nextPageUrl;
            }
            else
            {
                return null;
            }
        }

        private void GetPageFile(string pageUrl, Dictionary<string, string> listRow)
        {
            try
            {
                string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
                string filePath = this.RunPage.GetFilePath(pageUrl, pageSourceDir);
                if (!File.Exists(filePath))
                {
                    this.DownloadFile(pageUrl, filePath, listRow);
                }
                string nextPageUrl = this.GetNextPageUrl(pageUrl);
                if (nextPageUrl != null)
                {
                    this.GetPageFile(nextPageUrl, listRow);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        private void DownloadFile(string pageUrl, string filePath, Dictionary<string, string> listRow)
        {
            string text = this.RunPage.GetTextByRequest(pageUrl, listRow, false, 100, 30000, Encoding.UTF8, "", "", false, Proj_DataAccessType.WebRequestHtml, null, 100);
            FileHelper.SaveTextToFile(text, filePath, Encoding.UTF8); 
        }

        private ExcelWriter GetPageUrlsExcelWriter(int fileIndex)
        {
            String exportDir = this.RunPage.GetExportDir();
            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[] { 
                "detailPageUrl", 
                "detailPageName", 
                "cookie", 
                "grabStatus", 
                "giveUpGrab", 
                "keyword", 
                "品类", 
                "词类型", 
                "pageIndex" });

            string resultFilePath = Path.Combine(exportDir, "万方期刊_专业关键词_全部搜索页_" + fileIndex.ToString() + ".xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic);
            return resultEW;
        }

        private void GetAllPages(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            int fileIndex = 1;
            ExcelWriter ew = this.GetPageUrlsExcelWriter(fileIndex);
            Dictionary<string, string> idDic = new Dictionary<string, string>();
            int rowCount = 0;
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                if (rowCount % 1000 == 0)
                {
                    this.RunPage.InvokeAppendLogText("已处理到: fileIndex = " + fileIndex.ToString() + ", rowCount = " + rowCount.ToString(), LogLevelType.System, true);
                }

                if (rowCount >= 500000)
                {
                    if (ew != null)
                    {
                        ew.SaveToDisk();
                    }
                    fileIndex++;
                    ew = this.GetPageUrlsExcelWriter(fileIndex);
                    rowCount = 0;
                }

                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"];

                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string localFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir);

                    string nextPageUrl = detailUrl;
                    try
                    {
                        string pageFileText = FileHelper.GetTextFromFile(localFilePath);
                        HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                        htmlDoc.LoadHtml(pageFileText);
                        HtmlNode spanNode = htmlDoc.DocumentNode.SelectSingleNode("//span[@class=\"page_link\"]");
                        string keywordEncode = CommonUtil.StringToHexString(row["keyword"], Encoding.UTF8);
                        string detailPageName = row["keyword"] + "_" + row["品类"];

                        if (spanNode == null)
                        { 
                            Dictionary<string, string> f2vs = new Dictionary<string, string>();
                            f2vs.Add("detailPageUrl", nextPageUrl);
                            f2vs.Add("detailPageName", detailPageName + "_1");
                            f2vs.Add("keyword", row["keyword"]);
                            f2vs.Add("品类", row["品类"]);
                            f2vs.Add("词类型", row["词类型"]);
                            f2vs.Add("pageIndex", "1");
                            ew.AddRow(f2vs);
                            rowCount++;
                        }
                        else
                        {
                            int pageIndex = 1;

                            while (htmlDoc != null)
                            { 
                                Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                f2vs.Add("detailPageUrl", nextPageUrl);
                                f2vs.Add("detailPageName", detailPageName + "_" + pageIndex.ToString());
                                f2vs.Add("keyword", row["keyword"]);
                                f2vs.Add("品类", row["品类"]);
                                f2vs.Add("词类型", row["词类型"]);
                                f2vs.Add("pageIndex", pageIndex.ToString());
                                ew.AddRow(f2vs);
                                rowCount++;

                                HtmlNode nextPageNode = htmlDoc.DocumentNode.SelectSingleNode("//a[@class=\"page\"]/t[contains(text(),\"下一页\")]");
                                if (nextPageNode != null)
                                {
                                    nextPageUrl = "http://librarian.wanfangdata.com.cn/SearchResult.aspx" + CommonUtil.UrlDecodeSymbolAnd(nextPageNode.ParentNode.GetAttributeValue("href", ""));

                                    localFilePath = this.RunPage.GetFilePath(nextPageUrl, pageSourceDir);
                                    pageFileText = FileHelper.GetTextFromFile(localFilePath);
                                    htmlDoc = new HtmlAgilityPack.HtmlDocument();
                                    htmlDoc.LoadHtml(pageFileText);
                                    pageIndex++;
                                }
                                else
                                {
                                    htmlDoc = null;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        this.RunPage.InvokeAppendLogText(ex.Message + ". detailUrl = " + detailUrl, LogLevelType.Error, true);
                        throw ex;

                    }
                }
            }
            ew.SaveToDisk();
        } 
    }
}