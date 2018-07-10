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

namespace NetDataAccess.Extended.Celloud
{
    public class OalibGetFirstListPage : CustomProgramBase
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
            resultColumnDic.Add("species", 5); 
            resultColumnDic.Add("year", 6);
            string resultFilePath = Path.Combine(exportDir, "oalib获取列表页.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName; 

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                { 
                    string url = row[detailPageUrlColumnName];
                    string cookie = row["cookie"];
                    string species = row["species"].Trim(); 
                    string year = row["year"].Trim();
                    string localFilePath = this.RunPage.GetFilePath(url, pageSourceDir);
                    TextReader tr = null;

                    try
                    {
                        tr = new StreamReader(localFilePath);
                        string webPageHtml = tr.ReadToEnd();

                        HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                        htmlDoc.LoadHtml(webPageHtml);

                        HtmlNodeCollection allPageNodes = htmlDoc.DocumentNode.SelectNodes("//ul[@id=\"pages\"]/li");
                        if (allPageNodes != null)
                        {
                            HtmlNode lastPageNode = allPageNodes[allPageNodes.Count - 1];
                            string pageCountStr = lastPageNode.InnerText.Trim();
                            int pageCount = int.Parse(pageCountStr);
                            string searchKeyStr = CommonUtil.UrlEncode(species);
                            for (int j = 0; j < pageCount; j++)
                            {
                                string pageIndexStr = (j + 1).ToString();
                                string detailPageName = species + "_" + year + "_" + pageIndexStr;
                                string detailPageUrl = "http://www.oalib.com/search?type=0&oldType=0&kw=%22" + searchKeyStr + "%22&searchField=All&__multiselect_searchField=&fromYear=" + year + "&toYear=&pageNo=" + pageIndexStr;
                                Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                f2vs.Add("detailPageUrl", detailPageUrl);
                                f2vs.Add("detailPageName", detailPageName);
                                f2vs.Add("cookie", cookie);
                                f2vs.Add("species", species); 
                                f2vs.Add("year", year);
                                resultEW.AddRow(f2vs);
                            }
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