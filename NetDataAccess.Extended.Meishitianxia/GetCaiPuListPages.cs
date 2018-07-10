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

namespace NetDataAccess.Extended.Meishitianxia
{
    public class GetCaiPuListPages : ExternalRunWebPage
    { 
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            this.GetCategoryToPageUrls(listSheet);
            this.GetDetailPageUrls(listSheet);
            return true;
        }

        public override void GetDataByOtherAcessType(Dictionary<string, string> listRow)
        {
            string detailPageUrl = listRow[SysConfig.DetailPageUrlFieldName];
            string category = listRow["category"];
            string subCategory = listRow["subCategory"];
            string sourceDir = this.RunPage.GetDetailSourceFileDir();
            string subCategoryFilePath = this.RunPage.GetFilePath(detailPageUrl, sourceDir);
            if (!File.Exists(subCategoryFilePath))
            {
                string subCategoryDir = Path.Combine(sourceDir, subCategory);
                if (!Directory.Exists(subCategoryDir))
                {
                    Directory.CreateDirectory(subCategoryDir);
                }

                int pageCount = 0;
                bool needGetNextPage = true;
                while (needGetNextPage)
                {
                    int pageIndex = pageCount + 1;
                    string nextListPageUrl = this.GetNextListPageUrl(detailPageUrl, pageIndex);
                    string localPath = this.RunPage.GetFilePath(nextListPageUrl, subCategoryDir);
                    if (!File.Exists(localPath))
                    {
                        string pageHtml = this.RunPage.GetTextByRequest(nextListPageUrl, listRow, false, 0, 30000, Encoding.UTF8, null, null, false, Proj_DataAccessType.WebRequestHtml, null, 0);

                        HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                        htmlDoc.LoadHtml(pageHtml);

                        HtmlNode caipuListNode = htmlDoc.DocumentNode.SelectSingleNode("//div[@id=\"J_list\"]");
                        if (caipuListNode != null)
                        {
                            FileHelper.SaveTextToFile(pageHtml, localPath);
                            pageCount++;
                        }
                        else
                        {
                            needGetNextPage = true;
                        }
                    }
                    else
                    {
                        pageCount++;
                    }
                }

                this.SaveCaipuToSubCategoryFile(subCategoryFilePath, detailPageUrl, pageCount, subCategoryDir, category, subCategory);
            }
        }

        private void SaveCaipuToSubCategoryFile(string subCategoryFilePath, string detailPageUrl, int pageCount, string subCategoryDir, string category, string subCategory)
        {
            ExcelWriter subCategoryEW = this.CreateSubCategoryWriter(subCategoryFilePath);
            Dictionary<string, string> urlDic = new Dictionary<string, string>();
            for (int i = 0; i < pageCount; i++)
            {
                int pageIndex = i + 1;
                string nextListPageUrl = this.GetNextListPageUrl(detailPageUrl, pageIndex);
                string localPath = this.RunPage.GetFilePath(nextListPageUrl, subCategoryDir);
                string pageHtml = FileHelper.GetTextFromFile(localPath);
                HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                htmlDoc.LoadHtml(pageHtml);

                HtmlNodeCollection caipuNodeList = htmlDoc.DocumentNode.SelectNodes("//div[@id=\"J_list\"]/ul/li/div/h2/a");
                if (caipuNodeList != null)
                {
                    foreach (HtmlNode caipuNode in caipuNodeList)
                    {
                        string name = CommonUtil.HtmlDecode(caipuNode.InnerText).Trim();
                        string url = caipuNode.GetAttributeValue("href", "");

                        if (!urlDic.ContainsKey(url))
                        {
                            urlDic.Add(url, null);

                            Dictionary<string, string> f2vs = new Dictionary<string, string>(); 
                            f2vs.Add("name", name);
                            f2vs.Add("url", url);
                            f2vs.Add("category", category);
                            f2vs.Add("subCategory", subCategory);

                            subCategoryEW.AddRow(f2vs);
                        }
                    }
                }
            }
            subCategoryEW.SaveToDisk();
        }

        private string GetNextListPageUrl(string subCategoryUrl, int pageIndex)
        {
            if (subCategoryUrl.EndsWith(".html"))
            {
                return subCategoryUrl.Substring(0, subCategoryUrl.Length - 5) + "-page-" + pageIndex.ToString() + ".html";
            }
            else
            {
                return subCategoryUrl + "page/" + pageIndex.ToString() + "/";
            }
        }

        private ExcelWriter CreateSubCategoryWriter(string subCategoryFilePath)
        {
            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>(); 
            resultColumnDic.Add("name", 0);
            resultColumnDic.Add("url", 1);
            resultColumnDic.Add("category", 2);
            resultColumnDic.Add("subCategory", 3);
            ExcelWriter resultEW = new ExcelWriter(subCategoryFilePath, "List", resultColumnDic, null);
            return resultEW;
        }

        private ExcelWriter CreateSubCategoryMapWriter(string mapFilePath)
        {
            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("category", 0);
            resultColumnDic.Add("subCategory", 1);
            resultColumnDic.Add("name", 2);
            resultColumnDic.Add("url", 3);
            ExcelWriter resultEW = new ExcelWriter(mapFilePath, "List", resultColumnDic, null);
            return resultEW;
        }

        private ExcelWriter CreateDetailFileWriter(string subCategoryFilePath)
        {
            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("name", 5);
            resultColumnDic.Add("url", 6); 
            ExcelWriter resultEW = new ExcelWriter(subCategoryFilePath, "List", resultColumnDic, null);
            return resultEW;
        }


        private void GetCategoryToPageUrls(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            string resultFilePath = Path.Combine(exportDir, "美食天下_分类与菜谱列表对照.xlsx");

            ExcelWriter resultEW = this.CreateSubCategoryMapWriter(resultFilePath);

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailPageUrl = row[SysConfig.DetailPageUrlFieldName];
                string category = row["category"];
                string subCategory = row["subCategory"];
                string sourceDir = this.RunPage.GetDetailSourceFileDir();
                string subCategoryFilePath = this.RunPage.GetFilePath(detailPageUrl, sourceDir);

                ExcelReader er = new ExcelReader(subCategoryFilePath);
                int rowCount = er.GetRowCount();
                for (int j = 0; j < rowCount; j++)
                {
                    Dictionary<string, string> subRow = er.GetFieldValues(j);

                    Dictionary<string, string> mapRow = new Dictionary<string, string>();
                    mapRow.Add("category", subRow["category"]);
                    mapRow.Add("subCategory", subRow["subCategory"]);
                    mapRow.Add("name", subRow["name"]);
                    mapRow.Add("url", subRow["url"]);
                    resultEW.AddRow(mapRow);
                }
            }
            resultEW.SaveToDisk();
        }

        private void GetDetailPageUrls(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            string resultFilePath = Path.Combine(exportDir, "美食天下_获取菜谱详情页.xlsx");

            ExcelWriter resultEW = this.CreateDetailFileWriter(resultFilePath);
            Dictionary<string, string> urlDic = new Dictionary<string, string>();
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailPageUrl = row[SysConfig.DetailPageUrlFieldName];
                string category = row["category"];
                string subCategory = row["subCategory"];
                string sourceDir = this.RunPage.GetDetailSourceFileDir();
                string subCategoryFilePath = this.RunPage.GetFilePath(detailPageUrl, sourceDir);

                ExcelReader er = new ExcelReader(subCategoryFilePath);
                int rowCount = er.GetRowCount();
                for (int j = 0; j < rowCount; j++)
                {
                    Dictionary<string, string> subRow = er.GetFieldValues(j);
                    string url = subRow["url"];
                    if (!urlDic.ContainsKey(url))
                    {
                        Dictionary<string, string> mapRow = new Dictionary<string, string>();
                        mapRow.Add("detailPageUrl", url);
                        mapRow.Add("detailPageName", url);
                        mapRow.Add("name", subRow["name"]);
                        mapRow.Add("url", url); 
                        resultEW.AddRow(mapRow);
                    }
                }
            }
            resultEW.SaveToDisk();
        }
         
    }
}