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
    public class GetYiTangXiaoLei : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        { 
            this.GetXiaoLeiPageUrls(listSheet);
            return true;
        }

        private void GetXiaoLeiPageUrls(IListSheet listSheet)
        {
            string sourceDir = this.RunPage.GetDetailSourceFileDir();
            ExcelWriter resultEW = this.CreateResultWriter();
            Dictionary<string, bool> pageDic = new Dictionary<string, bool>();
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> listRow = listSheet.GetRow(i);

                bool giveUp = "Y".Equals(listRow[SysConfig.GiveUpGrabFieldName]);
                if (giveUp)
                {
                    string pageUrl = listRow["detailPageUrl"];
                    if (!pageDic.ContainsKey(pageUrl))
                    {
                        pageDic.Add(pageUrl, true);
                        Dictionary<string, string> resultRow = new Dictionary<string, string>();
                        resultRow.Add("detailPageUrl", listRow["detailPageUrl"]);
                        resultRow.Add("detailPageName", listRow["detailPageName"]);
                        resultRow.Add("giveUpGrab", "Y");
                        resultRow.Add("name", listRow["name"]);
                        resultRow.Add("pageType", listRow["pageType"]);
                        resultEW.AddRow(resultRow);
                    }
                }
                else
                {
                    HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);

                    HtmlNodeCollection linkNodes = htmlDoc.DocumentNode.SelectNodes("//ul[@class=\"ulLi120 fsc16\"]/li/a");
                    foreach (HtmlNode linkNode in linkNodes)
                    {
                        string linkUrl = linkNode.GetAttributeValue("href", "");
                        string fullLinkUrl = "http://www.yitang.org" + linkUrl;

                        if (!pageDic.ContainsKey(fullLinkUrl))
                        {
                            pageDic.Add(fullLinkUrl, true);
                            string linkName = CommonUtil.HtmlDecode(linkNode.InnerText).Trim();

                            Dictionary<string, string> resultRow = new Dictionary<string, string>();
                            resultRow.Add("detailPageUrl", fullLinkUrl);
                            resultRow.Add("detailPageName", linkName + "_word");
                            resultRow.Add("giveUpGrab", "Y");
                            resultRow.Add("name", linkName);
                            resultRow.Add("pageType", "word");
                            resultEW.AddRow(resultRow);
                        }
                    } 

                    HtmlNodeCollection pageLinkNodes = htmlDoc.DocumentNode.SelectNodes("//div[@class=\"pages\"]/a");
                    if (pageLinkNodes != null)
                    {
                        foreach (HtmlNode pageLinkNode in pageLinkNodes)
                        {
                            string pageLinkUrl = pageLinkNode.GetAttributeValue("href", "");
                            string fullPageLinkUrl = "http://www.yitang.org" + pageLinkUrl;

                            if (!pageDic.ContainsKey(fullPageLinkUrl))
                            {
                                pageDic.Add(fullPageLinkUrl, true);

                                Dictionary<string, string> resultRow = new Dictionary<string, string>();
                                resultRow.Add("detailPageUrl", fullPageLinkUrl);
                                resultRow.Add("detailPageName", pageLinkUrl);
                                resultRow.Add("giveUpGrab", "N");
                                resultRow.Add("name", pageLinkUrl);
                                resultRow.Add("pageType", "list");
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
            string resultFilePath = Path.Combine(exportDir, "成语_YiTang_小类列表页.xlsx");

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