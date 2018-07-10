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

namespace NetDataAccess.Extended.Yiguo
{
    public class SMListPageUrl : CustomProgramBase
    {
        public bool Run(string parameters, IListSheet listSheet)
        {
            return GetAllListPageUrl(listSheet);
        }
        private bool GetAllListPageUrl(IListSheet listSheet)
        {
            string exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            
            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("fullCategoryName", 5);
            resultColumnDic.Add("category1Code", 6);
            resultColumnDic.Add("category2Code", 7);
            resultColumnDic.Add("category3Code", 8);
            resultColumnDic.Add("category1Name", 9);
            resultColumnDic.Add("category2Name", 10);
            resultColumnDic.Add("category3Name", 11);
            resultColumnDic.Add("district", 12);
            string resultFilePath = Path.Combine(exportDir, this.RunPage.Project.Name + "_AllListPageUrl.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;
            string categoryNameColumnName = SysConfig.DetailPageNameFieldName;

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string fullCategoryName = row[categoryNameColumnName];
                    string url = row[detailPageUrlColumnName];
                    string cookie = row["cookie"];
                    string category1Code = row["category1Code"];
                    string category2Code = row["category2Code"];
                    string category3Code = row["category3Code"];
                    string category1Name = row["category1Name"];
                    string category2Name = row["category2Name"];
                    string category3Name = row["category3Name"];
                    string district = row["district"]; 
                    string localFilePath = this.RunPage.GetFilePath(url, pageSourceDir);
                    TextReader tr = null;

                    try
                    {
                        tr = new StreamReader(localFilePath,  Encoding.GetEncoding("GBK"));
                        string webPageHtml = tr.ReadToEnd();

                        HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                        htmlDoc.LoadHtml(webPageHtml);

                        HtmlNode pageNode = htmlDoc.DocumentNode.SelectSingleNode("//div[@class=\"pageNav fr\"]/ins");
                        string pageStr = pageNode.InnerText.Trim();
                        string[] pageSplits = pageStr.Split(new string[] { "/" }, StringSplitOptions.RemoveEmptyEntries);
                        string pageCountStr = pageSplits[1]; 
                        int pageCount = int.Parse(pageCountStr);
                        int dotIndex = url.LastIndexOf(".");
                        string urlPrefix = url.Substring(0, dotIndex);
                        string urlPostfix = url.Substring(dotIndex);
                        for (int j = 0; j < pageCount; j++)
                        {
                            string pageIndexStr = (j + 1).ToString();
                            string detailPageName = fullCategoryName + "_" + pageIndexStr;
                            string detailPageUrl = urlPrefix + "-" + (j + 1) + urlPostfix;
                            Dictionary<string, string> f2vs = new Dictionary<string, string>();
                            f2vs.Add("detailPageUrl", detailPageUrl);
                            f2vs.Add("detailPageName", detailPageName);
                            f2vs.Add("cookie", cookie);
                            f2vs.Add("fullCategoryName", fullCategoryName);
                            f2vs.Add("category1Code", category1Code);
                            f2vs.Add("category2Code", category2Code);
                            f2vs.Add("category3Code", category3Code);
                            f2vs.Add("category1Name", category1Name);
                            f2vs.Add("category2Name", category2Name);
                            f2vs.Add("category3Name", category3Name);
                            f2vs.Add("district", district);                             
                            resultEW.AddRow(f2vs);
                        }
                    }
                    catch (Exception ex)
                    {
                        if (tr != null)
                        {
                            tr.Dispose();
                            tr = null;
                        }
                        this.RunPage.InvokeAppendLogText("读取出错.  " + ex.Message + " LocalPath = " + localFilePath, LogLevelType.Error, true);
                        throw ex;
                    }
                }
            }
            resultEW.SaveToDisk();

            return true;
        }
    }
}