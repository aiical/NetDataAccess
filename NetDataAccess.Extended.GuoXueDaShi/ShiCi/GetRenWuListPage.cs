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

namespace NetDataAccess.Extended.GuoXueDaShi.ShiCi
{
    public class GetRenWuListPage : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        { 
            this.GetRenWuDetailPageUrls(listSheet);
            return true;
        }

        private void GetRenWuDetailPageUrls(IListSheet listSheet)
        {
            string sourceDir = this.RunPage.GetDetailSourceFileDir();
            ExcelWriter resultEW = this.CreateResultWriter();
            Dictionary<string, bool> pageUrlDic = new Dictionary<string, bool>();
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> listRow = listSheet.GetRow(i);
                HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);

                HtmlNodeCollection shiDaiNodes = htmlDoc.DocumentNode.SelectNodes("//div[@class=\"info_content zj clearfix\"]/dl");
                foreach (HtmlNode shiDaiNode in shiDaiNodes)
                {
                    HtmlNode shiDaiNameNode = shiDaiNode.SelectSingleNode("./dt");
                    string shiDaiName = CommonUtil.HtmlDecode(shiDaiNameNode.InnerText).Trim().Replace("【", "").Replace("】", "");

                    HtmlNodeCollection renWuNodes = shiDaiNode.SelectNodes("./dd/a");
                    if (renWuNodes != null)
                    {
                        foreach (HtmlNode renWuNode in renWuNodes)
                        {
                            string renWuPageUrl =  renWuNode.GetAttributeValue("href", "");
                            string fullRenWuPageUrl = "http://www.guoxuedashi.com" + renWuPageUrl;
                            if (!pageUrlDic.ContainsKey(fullRenWuPageUrl))
                            {
                                pageUrlDic.Add(fullRenWuPageUrl, true);
                                string renWu = CommonUtil.HtmlDecode(renWuNode.InnerText).Trim();
                                Dictionary<string, string> resultRow = new Dictionary<string, string>();
                                resultRow.Add("detailPageUrl", fullRenWuPageUrl);
                                resultRow.Add("detailPageName", fullRenWuPageUrl);
                                resultRow.Add("renWu", renWu);
                                resultRow.Add("shiDai", shiDaiName);
                                resultEW.AddRow(resultRow);
                            }
                        }
                    }
                }
            }

            resultEW.SaveToDisk();
        }

        private ExcelWriter CreateResultWriter()
        {
            String exportDir = this.RunPage.GetExportDir();
            string resultFilePath = Path.Combine(exportDir, "国学大师_诗词_人物页面.xlsx");

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("renWu", 5);
            resultColumnDic.Add("shiDai", 6); 
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }
    }
}