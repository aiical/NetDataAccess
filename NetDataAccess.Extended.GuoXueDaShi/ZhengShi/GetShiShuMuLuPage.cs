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

namespace NetDataAccess.Extended.GuoXueDaShi.ZhengShi
{
    public class GetShiShuMuLuPage : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        { 
            this.GetShiShuDetailPageUrls(listSheet);
            return true;
        }

        private void GetShiShuDetailPageUrls(IListSheet listSheet)
        {
            string sourceDir = this.RunPage.GetDetailSourceFileDir();
            ExcelWriter resultEW = this.CreateResultWriter(); 
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> listRow = listSheet.GetRow(i);
                HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);

                HtmlNodeCollection linkNodes = htmlDoc.DocumentNode.SelectNodes("//div[@class=\"info_cate clearfix\"]/dl/dd/a");
                foreach (HtmlNode linkNode in linkNodes)
                {

                    string juanName = CommonUtil.HtmlDecode(linkNode.InnerText).Trim();

                    string juanPageUrl = linkNode.GetAttributeValue("href", "");
                    string fullJuanPageUrl = "http://www.guoxuedashi.com" + juanPageUrl; 
                    Dictionary<string, string> resultRow = new Dictionary<string, string>();
                    resultRow.Add("detailPageUrl", fullJuanPageUrl);
                    resultRow.Add("detailPageName", fullJuanPageUrl);
                    resultRow.Add("shiShu", listRow["shiShu"]);
                    resultRow.Add("leiXing", listRow["leiXing"]);
                    resultRow.Add("juan", juanName);
                    resultEW.AddRow(resultRow);
                }
            }

            resultEW.SaveToDisk();
        }

        private ExcelWriter CreateResultWriter()
        {
            String exportDir = this.RunPage.GetExportDir();
            string resultFilePath = Path.Combine(exportDir, "国学大师_正史_卷页面.xlsx");

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("shiShu", 5);
            resultColumnDic.Add("leiXing", 6);
            resultColumnDic.Add("juan", 7); 
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }
    }
}