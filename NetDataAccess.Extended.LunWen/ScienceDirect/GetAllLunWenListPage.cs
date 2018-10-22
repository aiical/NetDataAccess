using HtmlAgilityPack;
using NetDataAccess.Base.Browser;
using NetDataAccess.Base.Common;
using NetDataAccess.Base.Config;
using NetDataAccess.Base.DB;
using NetDataAccess.Base.EnumTypes;
using NetDataAccess.Base.Reader;
using NetDataAccess.Base.UI;
using NetDataAccess.Base.Web;
using NetDataAccess.Base.Writer;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace NetDataAccess.Extended.LunWen.ScienceDirect
{ 
    public class GetAllLunWenListPage : ExternalRunWebPage
    { 
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            string exportDir = this.RunPage.GetExportDir();
            string exportFilePath = Path.Combine(exportDir, "论文_ScienceDirect_论文详情页.xlsx");
            ExcelWriter resultWriter = this.GetExcelWriter(exportFilePath);

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> listRow = listSheet.GetRow(i);
                string pageUrl = listRow[SysConfig.DetailPageUrlFieldName];
                string prefixUrl = this.GetUrlPrefix(pageUrl);
                String sourceDir = this.RunPage.GetDetailSourceFileDir();
                string sourceFilePath = this.RunPage.GetFilePath(pageUrl, sourceDir);

                HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                HtmlNodeCollection allLinkNodes = htmlDoc.DocumentNode.SelectNodes("//a[@class=\"anchor article-content-title u-margin-xs-top u-margin-s-bottom\"]");
                if (allLinkNodes != null)
                {
                    for (int j = 0; j < allLinkNodes.Count; j++)
                    {
                        HtmlNode linkNode = allLinkNodes[j];
                        string name = CommonUtil.HtmlDecode(linkNode.InnerText).Trim();
                        string url = prefixUrl + linkNode.GetAttributeValue("href", "");

                        Dictionary<string, string> row = new Dictionary<string, string>();
                        row.Add("detailPageUrl", url);
                        row.Add("detailPageName", url);
                        row.Add("name", name);
                        resultWriter.AddRow(row);
                    }
                } 
            }
            resultWriter.SaveToDisk();
            return true;
        }

        private string GetUrlPrefix(string listPageUrl)
        {
            int endIndex = listPageUrl.IndexOf("journal") - 1;
            string prefix = listPageUrl.Substring(0, endIndex);
            return prefix;
        }

        private ExcelWriter GetExcelWriter(string filePath)        
        { 
            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("name", 5);
            
            ExcelWriter resultEW = new ExcelWriter(filePath, "List", resultColumnDic, null);
            return resultEW;
        } 
    }
}
