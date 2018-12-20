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
    public class GetRenWuDetailPage : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            this.GetRenWuInfos(listSheet);
            this.GetShiCiDetailPageUrls(listSheet);
            return true;
        }

        private void GetRenWuInfos(IListSheet listSheet)
        {
            string sourceDir = this.RunPage.GetDetailSourceFileDir();
            ExcelWriter resultEW = this.CreateRenWuResultWriter();
            Dictionary<string, bool> pageUrlDic = new Dictionary<string, bool>();
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> listRow = listSheet.GetRow(i);
                HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);

                HtmlNode mainInfoNode = htmlDoc.DocumentNode.SelectSingleNode("//div[@class=\"info_txt2 clearfix\"]");
                HtmlNode titleNode = mainInfoNode.SelectSingleNode("./h2");
                string renWuTitle = CommonUtil.HtmlDecode(titleNode.InnerText).Trim();
                HtmlNode descriptionNode = mainInfoNode.SelectSingleNode("./p");
                string description = descriptionNode == null ? "" : CommonUtil.HtmlDecode(descriptionNode.InnerText).Trim();

                Dictionary<string, string> resultRow = new Dictionary<string, string>(); 
                resultRow.Add("人物", listRow["renWu"]);
                resultRow.Add("时代", listRow["shiDai"]);
                resultRow.Add("人物页面标题", renWuTitle);
                resultRow.Add("简介", description);
                resultRow.Add("url", listRow[SysConfig.DetailPageUrlFieldName]);
                resultEW.AddRow(resultRow);
            }

            resultEW.SaveToDisk();
        }

        private ExcelWriter CreateRenWuResultWriter()
        {
            String exportDir = this.RunPage.GetExportDir();
            string resultFilePath = Path.Combine(exportDir, "国学大师_诗词_人物信息.xlsx");

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("人物", 0);
            resultColumnDic.Add("时代", 1);
            resultColumnDic.Add("人物页面标题", 2);
            resultColumnDic.Add("简介", 3);
            resultColumnDic.Add("url", 4);
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }

        private void GetShiCiDetailPageUrls(IListSheet listSheet)
        {
            string sourceDir = this.RunPage.GetDetailSourceFileDir();
            ExcelWriter resultEW = this.CreateShiCiUrlResultWriter();
            Dictionary<string, bool> pageUrlDic = new Dictionary<string, bool>();
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> listRow = listSheet.GetRow(i);
                HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);

                string renWu = listRow["renWu"];
                string shiDai = listRow["shiDai"];

                HtmlNodeCollection shiCiNodes = htmlDoc.DocumentNode.SelectNodes("//div[@class=\"info_cate clearfix\"]/dl/dd/a");
                foreach (HtmlNode shiCiNode in shiCiNodes)
                {
                    string shiCiName = CommonUtil.HtmlDecode(shiCiNode.InnerText).Trim();

                    string shiCiDetailPageUrl = shiCiNode.GetAttributeValue("href", "");
                    string fullShiCiDetailPageUrl = "http://www.guoxuedashi.com" + shiCiDetailPageUrl;
                    if (!pageUrlDic.ContainsKey(fullShiCiDetailPageUrl))
                    {
                        pageUrlDic.Add(fullShiCiDetailPageUrl, true); 
                        Dictionary<string, string> resultRow = new Dictionary<string, string>();
                        resultRow.Add("detailPageUrl", fullShiCiDetailPageUrl);
                        resultRow.Add("detailPageName", fullShiCiDetailPageUrl);
                        resultRow.Add("renWu", renWu);
                        resultRow.Add("shiDai", shiDai);
                        resultRow.Add("shiCi", shiCiName); 
                        resultEW.AddRow(resultRow);
                    }
                }
            }

            resultEW.SaveToDisk();
        }

        private ExcelWriter CreateShiCiUrlResultWriter()
        {
            String exportDir = this.RunPage.GetExportDir();
            string resultFilePath = Path.Combine(exportDir, "国学大师_诗词_诗词页面.xlsx");

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("renWu", 5);
            resultColumnDic.Add("shiDai", 6);
            resultColumnDic.Add("shiCi", 7); 
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }
    }
}