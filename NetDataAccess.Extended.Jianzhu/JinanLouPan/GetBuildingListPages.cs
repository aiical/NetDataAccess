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

namespace NetDataAccess.Extended.Jianzhu.JinanLouPan
{
    public class GetBuildingListPages : ExternalRunWebPage
    {
        public override void CheckRequestCompleteFile(string webPageText, Dictionary<string, string> listRow)
        {
            if (webPageText.Contains("济南市城乡建设委员会") && webPageText.Trim().EndsWith("</html>"))
            {
            }
            else
            {
                throw new Exception("未能完整获取页面");
            }
        }

        public override bool AfterAllGrab(IListSheet listSheet)
        {
            try
            {
                return this.GetBuildingPageUrls(listSheet) && this.GetHouseListPageUrls(listSheet);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        private bool GetBuildingPageUrls(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("projectId", 5);
            resultColumnDic.Add("projectName", 6);
            resultColumnDic.Add("buildingId", 7);
            resultColumnDic.Add("楼名称", 8);
            resultColumnDic.Add("预售许可证", 9);
            resultColumnDic.Add("总套数", 10);
            resultColumnDic.Add("总面积", 11);
            resultColumnDic.Add("可售套数", 12);
            resultColumnDic.Add("可售面积", 13);
            resultColumnDic.Add("已售套数", 14);
            resultColumnDic.Add("已售面积", 15);
            string resultFilePath = Path.Combine(exportDir, "济南楼盘_楼详情页.xlsx");
            Dictionary<string, string> resultColumnFormat = new Dictionary<string, string>();
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, resultColumnFormat);

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;

            Dictionary<string, string> buildingDic = new Dictionary<string, string>();

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string url = row[detailPageUrlColumnName];
                    string projectId = row["projectId"];
                    string projectName = row["projectName"];

                    HtmlAgilityPack.HtmlDocument pageHtmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                    HtmlNodeCollection buildingNodeList = pageHtmlDoc.DocumentNode.SelectNodes("//table[@class=\"project_table\"]/tr");
                    for (int j = 1; j < buildingNodeList.Count - 1; j++)
                    {
                        HtmlNodeCollection buildingFieldNodeList = buildingNodeList[j].SelectNodes("./td");
                        HtmlNode buildingNameNode = buildingFieldNodeList[1];
                        string buildingName = buildingNameNode.GetAttributeValue("title", "");
                        string buildingUrl = buildingNameNode.SelectSingleNode("./a").GetAttributeValue("href", "");
                        int equalIndex = buildingUrl.LastIndexOf("=");
                        string buildingId = buildingUrl.Substring(equalIndex + 1).Trim();
                        if (!buildingDic.ContainsKey(buildingId))
                        {
                            buildingDic.Add(buildingId, null);
                            string detailPageUrl = "http://www.jnfdc.gov.cn/onsaling/bshow.shtml?bno=" + buildingId;
                            Dictionary<string, object> f2vs = new Dictionary<string, object>();
                            f2vs.Add("detailPageUrl", detailPageUrl);
                            f2vs.Add("detailPageName", buildingId);
                            f2vs.Add("projectId", projectId);
                            f2vs.Add("projectName", projectName);
                            f2vs.Add("buildingId", buildingId);
                            f2vs.Add("楼名称", buildingName);
                            f2vs.Add("预售许可证", buildingFieldNodeList[2].InnerText.Trim());
                            f2vs.Add("总套数", buildingFieldNodeList[3].InnerText.Trim());
                            f2vs.Add("总面积", buildingFieldNodeList[4].InnerText.Trim());
                            f2vs.Add("可售套数", buildingFieldNodeList[5].InnerText.Trim());
                            f2vs.Add("可售面积", buildingFieldNodeList[6].InnerText.Trim());
                            f2vs.Add("已售套数", buildingFieldNodeList[7].InnerText.Trim());
                            f2vs.Add("已售面积", buildingFieldNodeList[8].InnerText.Trim());
                            resultEW.AddRow(f2vs);
                        }
                    }
                }
            }

            resultEW.SaveToDisk();

            return true;
        }
        private bool GetHouseListPageUrls(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("projectId", 5);
            resultColumnDic.Add("projectName", 6);
            resultColumnDic.Add("buildingId", 7);
            resultColumnDic.Add("楼名称", 8); 
            string resultFilePath = Path.Combine(exportDir, "济南楼盘_房间列表页.xlsx");
            Dictionary<string, string> resultColumnFormat = new Dictionary<string, string>();
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, resultColumnFormat);

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;

            Dictionary<string, string> buildingDic = new Dictionary<string, string>();

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string url = row[detailPageUrlColumnName];
                    string projectId = row["projectId"];
                    string projectName = row["projectName"];

                    HtmlAgilityPack.HtmlDocument pageHtmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                    HtmlNodeCollection buildingNodeList = pageHtmlDoc.DocumentNode.SelectNodes("//table[@class=\"project_table\"]/tr");
                    for (int j = 1; j < buildingNodeList.Count - 1; j++)
                    {
                        HtmlNodeCollection buildingFieldNodeList = buildingNodeList[j].SelectNodes("./td");
                        HtmlNode buildingNameNode = buildingFieldNodeList[1];
                        string buildingName = buildingNameNode.GetAttributeValue("title", "");
                        string buildingUrl = buildingNameNode.SelectSingleNode("./a").GetAttributeValue("href", "");
                        int equalIndex = buildingUrl.LastIndexOf("=");
                        string buildingId = buildingUrl.Substring(equalIndex + 1).Trim();
                        if (!buildingDic.ContainsKey(buildingId))
                        {
                            buildingDic.Add(buildingId, null);
                            string detailPageUrl = "http://www.jnfdc.gov.cn/onsaling/viewhouse.shtml?fmid=" + buildingId;
                            Dictionary<string, object> f2vs = new Dictionary<string, object>();
                            f2vs.Add("detailPageUrl", detailPageUrl);
                            f2vs.Add("detailPageName", buildingId);
                            f2vs.Add("projectId", projectId);
                            f2vs.Add("projectName", projectName);
                            f2vs.Add("buildingId", buildingId);
                            f2vs.Add("楼名称", buildingName); 
                            resultEW.AddRow(f2vs);
                        }
                    }
                }
            }

            resultEW.SaveToDisk();

            return true;
        }
    }
}