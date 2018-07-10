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

namespace NetDataAccess.Extended.Dzdp
{
    public class GetFirstListPage : ExternalRunWebPage
    {
        private ExcelWriter GetExcelWriter(int fileIndex)
        {
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("city", 5);
            resultColumnDic.Add("g", 6);
            resultColumnDic.Add("r", 7);
            resultColumnDic.Add("gName", 8);
            resultColumnDic.Add("rName", 9);
            resultColumnDic.Add("pageIndex", 10);
            string resultFilePath = Path.Combine(exportDir, "大众点评获取列表页_" + fileIndex.ToString() + ".xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }

        public override bool AfterAllGrab(IListSheet listSheet)
        {
            ExcelWriter resultEW = null;
            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;
            int fileIndex = 0; 
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                if (resultEW == null || resultEW.RowCount > 200000)
                {
                    if (resultEW != null)
                    {
                        resultEW.SaveToDisk();
                    }
                    resultEW = this.GetExcelWriter(fileIndex);
                    fileIndex++;
                }

                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string url = row[detailPageUrlColumnName];
                    string cookie = row["cookie"];
                    string city = row["city"];
                    string g = row["g"];
                    string r = row["r"];
                    string gName = row["gName"];
                    string rName = row["rName"];

                    HtmlAgilityPack.HtmlDocument pageHtmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                    HtmlNodeCollection allPageLinkNodes = pageHtmlDoc.DocumentNode.SelectNodes("//div[@class=\"page\"]/a");

                    int pageCount = 1;
                    if (allPageLinkNodes != null)
                    {
                        HtmlNode lastPageLinkNode = allPageLinkNodes[allPageLinkNodes.Count - 2];
                        pageCount = int.Parse(lastPageLinkNode.InnerText);
                    }
                    for (int j = 0; j < pageCount; j++)
                    {
                        string pageIndex = "p" + (j + 1).ToString();
                        string detailPageName = city + "_" + g + r + pageIndex;
                        string detailPageUrl = url + (j == 0 ? "" : pageIndex);
                        Dictionary<string, string> f2vs = new Dictionary<string, string>();
                        f2vs.Add("detailPageUrl", detailPageUrl);
                        f2vs.Add("detailPageName", detailPageName);
                        f2vs.Add("cookie", cookie);
                        f2vs.Add("city", city);
                        f2vs.Add("g", g);
                        f2vs.Add("r", r);
                        f2vs.Add("gName", gName);
                        f2vs.Add("rName", rName);
                        f2vs.Add("pageIndex", pageIndex);
                        resultEW.AddRow(f2vs);
                    }
                }
            }

            resultEW.SaveToDisk();

            return true;
        }
        private void SaveRow(string infoName, string nodeHref, Dictionary<string, string> code2Names)
        {
            string[] infoPieces = nodeHref.Split(new string[] { "/" }, StringSplitOptions.RemoveEmptyEntries);
            string infoValue = infoPieces[infoPieces.Length - 1];
            if (!code2Names.ContainsKey(infoValue))
            {
                code2Names.Add(infoValue, infoName);
            }
        }
    }
}