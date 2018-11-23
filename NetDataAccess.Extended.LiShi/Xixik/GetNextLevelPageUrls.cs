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

namespace NetDataAccess.Extended.LiShi.Xixik
{
    public class GetNextLevelPageUrls : ExternalRunWebPage
    { 
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            this.GetNextPageUrls(listSheet); 
            return true;
        }

        private void GetNextPageUrls(IListSheet listSheet)
        {
            ExcelWriter resultEw = this.CreateNextPageUrlExcelWriter();
            Dictionary<string,bool> urlDic = new Dictionary<string,bool>();
            int rowCount = listSheet.RowCount;
            for (int i = 0; i < rowCount; i++)
            {
                Dictionary<string, string> listRow = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(listRow[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string pageUrl = listRow[SysConfig.DetailPageUrlFieldName];

                    if (!urlDic.ContainsKey(pageUrl))
                    {
                        Dictionary<string, string> oldRow = new Dictionary<string, string>();
                        oldRow.Add(SysConfig.DetailPageUrlFieldName, pageUrl);
                        oldRow.Add(SysConfig.DetailPageNameFieldName, pageUrl);
                        oldRow.Add("linkName", listRow["linkName"]);
                        resultEw.AddRow(oldRow);
                        urlDic.Add(pageUrl, true);
                    }

                    HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i, Encoding.GetEncoding("gb2312"));
                    HtmlNodeCollection linkNodes = htmlDoc.DocumentNode.SelectNodes("//a");
                    if (linkNodes != null)
                    {
                        foreach (HtmlNode linkNode in linkNodes)
                        {
                            string linkUrl = linkNode.GetAttributeValue("href", "").Trim();
                            if (linkUrl.StartsWith("http://114.xixik.com/") && !urlDic.ContainsKey(linkUrl))
                            {
                                string linkText = CommonUtil.HtmlDecode(linkNode.InnerText).Trim();
                                Dictionary<string, string> newRow = new Dictionary<string, string>();
                                newRow.Add(SysConfig.DetailPageUrlFieldName, linkUrl);
                                newRow.Add(SysConfig.DetailPageNameFieldName, linkUrl);
                                newRow.Add("linkName", linkText);
                                resultEw.AddRow(newRow);

                                urlDic.Add(linkUrl, true);
                            }
                        }
                    }

                }
            }
            resultEw.SaveToDisk();
        }
        private ExcelWriter CreateNextPageUrlExcelWriter()
        {
            String exportDir = this.RunPage.GetExportDir();
            string resultFilePath = Path.Combine(exportDir, "中国历史信息_xixik.xlsx");

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("linkName", 5);
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }
         
    }
}