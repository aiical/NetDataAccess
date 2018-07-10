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
    public class GetYearReportListPageA : ExternalRunWebPage
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
            string resultFilePath = Path.Combine(exportDir, "统计年鉴_山东_详情页A.xlsx"); 

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
                    HtmlNode mainUlNode = pageHtmlDoc.DocumentNode.SelectSingleNode("//body/center/div/table/tbody/center/tr/th/ul");

                    if (mainUlNode != null)
                    {
                        HtmlNodeCollection childNodes = mainUlNode.ChildNodes;
                        this.GetChildNodeInfos(resultEW, childNodes, dirUrl, "", year);
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
        private void GetChildNodeInfos(ExcelWriter resultEW, HtmlNodeCollection allNodes, string dirUrl, string parentCode, string year)
        {
            if (allNodes != null)
            {
                string childCode = "";
                int index = 0;
                for (int j = 0; j < allNodes.Count; j++)
                {
                    HtmlNode node = allNodes[j];
                    string nodeName = node.Name.ToLower().Trim();
                    if (nodeName == "ul")
                    {
                        GetChildNodeInfos(resultEW, node.ChildNodes, dirUrl, childCode, year);
                    }
                    else
                    {

                        string detailPageUrl = "";
                        string name = CommonUtil.HtmlDecode(node.InnerText).Trim();
                        bool giveUpGrab = true;
                        bool needSave = false;
                        if (nodeName == "a")
                        {
                            detailPageUrl = dirUrl + node.GetAttributeValue("href", "").Trim();
                            if (detailPageUrl.EndsWith(".xls"))
                            {
                                giveUpGrab = false;
                            }
                            needSave = true; 
                        }
                        else if (nodeName == "li")
                        {
                            HtmlNode linkNode = node.SelectSingleNode("./a");
                            if (linkNode != null)
                            {
                                detailPageUrl = dirUrl + linkNode.GetAttributeValue("href", "").Trim();
                                if (detailPageUrl.EndsWith(".xls"))
                                {
                                    giveUpGrab = false;
                                }
                            }
                            needSave = true;
                        }

                        if (needSave)
                        {
                            childCode = parentCode + index.ToString().PadLeft(2, '0');
                            index++;

                            if (detailPageUrl.Length == 0)
                            {
                                detailPageUrl = childCode;
                            }

                            Dictionary<string, object> f2vs = new Dictionary<string, object>();
                            f2vs.Add("detailPageUrl", detailPageUrl);
                            f2vs.Add("detailPageName", detailPageUrl);
                            f2vs.Add("giveUpGrab", giveUpGrab ? "Y" : "N");
                            f2vs.Add("code", childCode);
                            f2vs.Add("name", name);
                            f2vs.Add("year", year);
                            resultEW.AddRow(f2vs);
                        }
                    }
                }
            }
        }
    }
}