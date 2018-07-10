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
using System.Web;

namespace NetDataAccess.Extended.Tongjiju.Shandong
{
    public class GetYearReportListPageB : ExternalRunWebPage
    { 
        public ExcelWriter GetExcelWriter()
        {
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("code", 5);
            resultColumnDic.Add("name", 6);
            resultColumnDic.Add("year", 7);
            string resultFilePath = Path.Combine(exportDir, "统计年鉴_山东_html详情页B.xlsx"); 

            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }

        public override bool AfterAllGrab(IListSheet listSheet)
        {
            ExcelWriter resultEW = this.GetExcelWriter(); ; 

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;

            Dictionary<string, string> shopDic = new Dictionary<string, string>();

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string url = row[detailPageUrlColumnName];
                    Uri uri = new Uri(url);
                    string dirUrl = row["dirUrl"];
                    string year = row["year"];

                    HtmlAgilityPack.HtmlDocument pageHtmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i, Encoding.GetEncoding("gb2312"));
                    HtmlNodeCollection allDivNodes = pageHtmlDoc.DocumentNode.SelectNodes("//body/div/table/tbody/tr/th/div");

                    if (allDivNodes != null)
                    {
                        this.GetChildNodeInfos(resultEW, allDivNodes, dirUrl, year);
                    }
                    else
                    {
                        throw new Exception("找不到主ul");
                    }
                }
            }

            resultEW.SaveToDisk();

            return true;
        }
        private void GetChildNodeInfos(ExcelWriter resultEW, HtmlNodeCollection allNodes, string dirUrl, string year)
        {
            if (allNodes != null)
            { 
                for (int j = 0; j < allNodes.Count; j = j + 2)
                {
                    string categoryCode = j.ToString().PadLeft(2, '0');
                    HtmlNode categoryDivNode = allNodes[j];
                    string categoryName = CommonUtil.HtmlDecode(categoryDivNode.InnerText).Trim();
                    {
                        Dictionary<string, object> f2vs = new Dictionary<string, object>();
                        f2vs.Add("detailPageUrl", categoryName);
                        f2vs.Add("detailPageName", categoryName);
                        f2vs.Add("giveUpGrab", "Y");
                        f2vs.Add("code", categoryCode);
                        f2vs.Add("name", categoryName);
                        f2vs.Add("year", year);
                        resultEW.AddRow(f2vs);
                    }


                    HtmlNode listDivNode = allNodes[j + 1];
                    HtmlNodeCollection linkNodes = listDivNode.SelectNodes("./li/p/a");
                    if (linkNodes == null)
                    {
                        linkNodes = listDivNode.SelectNodes("./li/a");
                    }
                    for (int k = 0; k < linkNodes.Count; k++)
                    {
                        HtmlNode linkNode = linkNodes[k];
                        string detailPageUrl = dirUrl + linkNode.GetAttributeValue("href", "");
                        string name = CommonUtil.HtmlDecode(linkNode.InnerText).Trim();
                        string itemCode = categoryCode + k.ToString().PadLeft(2, '0');
                        Dictionary<string, object> f2vs = new Dictionary<string, object>();
                        f2vs.Add("detailPageUrl", detailPageUrl);
                        f2vs.Add("detailPageName", detailPageUrl); 
                        f2vs.Add("code", itemCode);
                        f2vs.Add("name", name);
                        f2vs.Add("year", year);
                        resultEW.AddRow(f2vs);
                    }
                }
            }
        }
    }
}