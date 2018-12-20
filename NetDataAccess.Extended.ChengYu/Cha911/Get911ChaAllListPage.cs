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

namespace NetDataAccess.Extended.ChengYu.Cha911
{
    public class Get911ChaAllListPage : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        { 
            this.GetDetailPageUrls(listSheet);
            return true;
        }

        private void GetDetailPageUrls(IListSheet listSheet)
        {
            string sourceDir = this.RunPage.GetDetailSourceFileDir();
            ExcelWriter resultEW = this.CreateResultWriter();
            Dictionary<string, bool> pageUrlDic = new Dictionary<string, bool>();
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> listRow = listSheet.GetRow(i);
                HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);

                HtmlNodeCollection pageUrlNodes = htmlDoc.DocumentNode.SelectNodes("//ul[@class=\"l5 center f14\"]/li/a");
                foreach (HtmlNode pageUrlNode in pageUrlNodes)
                {
                    string pageUrl = pageUrlNode.GetAttributeValue("href", "");
                    string fullPageUrl = "https://chengyu.911cha.com/" + pageUrl;
                    if (!pageUrlDic.ContainsKey(fullPageUrl))
                    {
                        string linkName = CommonUtil.HtmlDecode(pageUrlNode.InnerText).Trim(); 
                        pageUrlDic.Add(fullPageUrl, true);
                        Dictionary<string, string> resultRow = new Dictionary<string, string>();
                        resultRow.Add("detailPageUrl", fullPageUrl);
                        resultRow.Add("detailPageName", fullPageUrl);
                        resultEW.AddRow(resultRow);
                    }
                } 
            }

            resultEW.SaveToDisk();
        }

        private ExcelWriter CreateResultWriter()
        {
            String exportDir = this.RunPage.GetExportDir();
            string resultFilePath = Path.Combine(exportDir, "成语_911Cha_详情页.xlsx");

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4); 
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }
    }
}