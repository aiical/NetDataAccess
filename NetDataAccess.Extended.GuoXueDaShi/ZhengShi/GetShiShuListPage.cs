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
    public class GetShiShuListPage : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        { 
            this.GetShiShuMuLuPageUrls(listSheet);
            return true;
        }

        private void GetShiShuMuLuPageUrls(IListSheet listSheet)
        {
            string sourceDir = this.RunPage.GetDetailSourceFileDir();
            ExcelWriter resultEW = this.CreateResultWriter(); 
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> listRow = listSheet.GetRow(i);
                HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);

                HtmlNodeCollection dtNodes = htmlDoc.DocumentNode.SelectNodes("//div[@class=\"info_content clearfix\"]/dl/dt");
                foreach (HtmlNode dtNode in dtNodes)
                {
                    
                    string leiXingName = CommonUtil.HtmlDecode(dtNode.InnerText).Trim().Replace("【", "").Replace("】", "");

                    HtmlNode nextNode = dtNode.NextSibling;
                    while (nextNode != null && nextNode.Name.ToLower() != "dt")
                    {
                        if (nextNode.Name.ToLower() == "dd")
                        {
                            HtmlNode linkNode = nextNode.SelectSingleNode("./a");
                            if (linkNode != null)
                            {
                                string muLuPageUrl = linkNode.GetAttributeValue("href", "");
                                string fullMuLuPageUrl = muLuPageUrl.StartsWith("http") ? muLuPageUrl : ("http://www.guoxuedashi.com" + muLuPageUrl);
                                string shiShu = CommonUtil.HtmlDecode(linkNode.InnerText).Trim();
                                Dictionary<string, string> resultRow = new Dictionary<string, string>();
                                resultRow.Add("detailPageUrl", fullMuLuPageUrl);
                                resultRow.Add("detailPageName", fullMuLuPageUrl);
                                resultRow.Add("shiShu", shiShu);
                                resultRow.Add("leiXing", leiXingName);
                                resultEW.AddRow(resultRow);
                            }
                        }
                        nextNode = nextNode.NextSibling;
                    } 
                }
            }

            resultEW.SaveToDisk();
        }

        private ExcelWriter CreateResultWriter()
        {
            String exportDir = this.RunPage.GetExportDir();
            string resultFilePath = Path.Combine(exportDir, "国学大师_正史_目录页面.xlsx");

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("shiShu", 5);
            resultColumnDic.Add("leiXing", 6); 
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }
    }
}