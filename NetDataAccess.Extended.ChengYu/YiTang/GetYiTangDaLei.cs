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

namespace NetDataAccess.Extended.ChengYu.YiTang
{
    public class GetYiTangDaLei : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        { 
            this.GetXiaoLeiFirstPageUrls(listSheet);
            return true;
        }

        private void GetXiaoLeiFirstPageUrls(IListSheet listSheet)
        {
            string sourceDir = this.RunPage.GetDetailSourceFileDir();
            ExcelWriter resultEW = this.CreateResultWriter();
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> listRow = listSheet.GetRow(i);
                HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);

                HtmlNodeCollection linkNodes = htmlDoc.DocumentNode.SelectNodes("//ul[@class=\"ulLi120 fsc16\"]/li/a");
                foreach (HtmlNode linkNode in linkNodes)
                {
                    string linkUrl = linkNode.GetAttributeValue("href", "");
                    string linkName = CommonUtil.HtmlDecode(linkNode.InnerText).Trim();
                    string fullLinkUrl = "http://www.yitang.org" + linkUrl;

                    Dictionary<string, string> resultRow = new Dictionary<string, string>();
                    resultRow.Add("detailPageUrl", fullLinkUrl);
                    resultRow.Add("detailPageName", linkName + "_word");
                    resultRow.Add("giveUpGrab", "Y");
                    resultRow.Add("name", linkName);
                    resultRow.Add("pageType", "word");
                    resultEW.AddRow(resultRow);
                }

                HtmlNodeCollection moreLinkNodes = htmlDoc.DocumentNode.SelectNodes("//a[@class=\"more\"]");
                foreach (HtmlNode moreLinkNode in moreLinkNodes)
                {
                    string moreLinkUrl = moreLinkNode.GetAttributeValue("href", "");
                    string fullMoreLinkUrl = "http://www.yitang.org" + moreLinkUrl;

                    Dictionary<string, string> resultRow = new Dictionary<string, string>();
                    resultRow.Add("detailPageUrl", fullMoreLinkUrl);
                    resultRow.Add("detailPageName", moreLinkUrl);
                    resultRow.Add("giveUpGrab", "N");
                    resultRow.Add("name", moreLinkUrl);
                    resultRow.Add("pageType", "list");
                    resultEW.AddRow(resultRow);
                }
            }

            resultEW.SaveToDisk();
        }

        private ExcelWriter CreateResultWriter()
        {
            String exportDir = this.RunPage.GetExportDir();
            string resultFilePath = Path.Combine(exportDir, "成语_YiTang_小类列表页首页.xlsx");

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("name", 5);
            resultColumnDic.Add("pageType", 6); 
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }
    }
}