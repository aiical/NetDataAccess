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
    public class GetBuildingPages : ExternalRunWebPage
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
                return this.GetBuildingInfos(listSheet) && this.GetBuildingUseInfos(listSheet);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        private bool GetBuildingInfos(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("projectId", 0);
            resultColumnDic.Add("项目名称", 1);
            resultColumnDic.Add("buildingId", 2);
            resultColumnDic.Add("楼名称", 3);
            resultColumnDic.Add("总套数", 4);
            resultColumnDic.Add("总面积", 5);
            resultColumnDic.Add("可售套数", 6);
            resultColumnDic.Add("可售面积", 7);
            resultColumnDic.Add("已售套数", 8);
            resultColumnDic.Add("已售面积", 9);
            resultColumnDic.Add("开发单位", 10);
            resultColumnDic.Add("建筑面积（万平方米）", 11);
            resultColumnDic.Add("装修标准", 12);
            resultColumnDic.Add("规划用途", 13);
            resultColumnDic.Add("有无抵押", 14);
            resultColumnDic.Add("商品房预售许可证", 15);
            resultColumnDic.Add("国有土地使用证", 16);
            resultColumnDic.Add("建设工程规划许可证", 17);
            resultColumnDic.Add("建设工程施工许可证", 18);
            string resultFilePath = Path.Combine(exportDir, "济南楼盘_楼详情.xlsx");
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
                    string buildingId = row["buildingId"];
                    string buildingName = row["楼名称"];

                    HtmlAgilityPack.HtmlDocument pageHtmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                    HtmlNodeCollection buildingFieldNodeList = pageHtmlDoc.DocumentNode.SelectNodes("//table[@class=\"message_table\"]/tr");
                    HtmlNodeCollection buildingOtherFieldNodeList = buildingFieldNodeList[3].SelectNodes("./tr"); 
                    Dictionary<string, object> f2vs = new Dictionary<string, object>(); 
                    f2vs.Add("projectId", projectId);
                    f2vs.Add("项目名称", projectName);
                    f2vs.Add("buildingId", buildingId);
                    f2vs.Add("楼名称", buildingName);
                    f2vs.Add("总套数", row["总套数"]);
                    f2vs.Add("总面积", row["总面积"]);
                    f2vs.Add("可售套数", row["可售套数"]);
                    f2vs.Add("可售面积", row["可售面积"]);
                    f2vs.Add("已售套数", row["已售套数"]);
                    f2vs.Add("已售面积", row["已售面积"]);
                    f2vs.Add("开发单位", buildingFieldNodeList[2].SelectNodes("./td")[1].InnerText.Trim());
                    f2vs.Add("建筑面积（万平方米）", buildingOtherFieldNodeList[0].SelectNodes("./td")[3].InnerText.Trim());
                    f2vs.Add("装修标准", buildingOtherFieldNodeList[1].SelectNodes("./td")[3].InnerText.Trim());
                    f2vs.Add("规划用途", buildingOtherFieldNodeList[2].SelectNodes("./td")[1].InnerText.Trim());
                    f2vs.Add("有无抵押", buildingOtherFieldNodeList[2].SelectNodes("./td")[3].InnerText.Trim());
                    f2vs.Add("商品房预售许可证", buildingOtherFieldNodeList[3].SelectNodes("./td")[1].InnerText.Trim());
                    f2vs.Add("国有土地使用证", buildingOtherFieldNodeList[3].SelectNodes("./td")[3].InnerText.Trim());
                    f2vs.Add("建设工程规划许可证", buildingOtherFieldNodeList[4].SelectNodes("./td")[1].InnerText.Trim());
                    f2vs.Add("建设工程施工许可证", buildingOtherFieldNodeList[4].SelectNodes("./td")[3].InnerText.Trim());
                    resultEW.AddRow(f2vs);
                }
            }

            resultEW.SaveToDisk();

            return true;
        }

        private bool GetBuildingUseInfos(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("projectId", 0);
            resultColumnDic.Add("项目名称", 1);
            resultColumnDic.Add("buildingId", 2);
            resultColumnDic.Add("楼名称", 3);
            resultColumnDic.Add("用途", 4);
            resultColumnDic.Add("批准销售套数", 5);
            resultColumnDic.Add("批准销售面积", 6);
            resultColumnDic.Add("已售套数", 7);
            resultColumnDic.Add("已售面积", 8);
            resultColumnDic.Add("可售套数", 9);
            resultColumnDic.Add("可售面积", 10); 
            string resultFilePath = Path.Combine(exportDir, "济南楼盘_楼用途.xlsx");
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
                    string buildingId = row["buildingId"];
                    string buildingName = row["楼名称"];

                    HtmlAgilityPack.HtmlDocument pageHtmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                    HtmlNodeCollection buildingUseNodeList = pageHtmlDoc.DocumentNode.SelectNodes("//table[@class=\"flow_message_table\"]/tr");
                    for (int j = 1; j < buildingUseNodeList.Count; j++)
                    {
                        HtmlNode buildingUseNode = buildingUseNodeList[j];
                        HtmlNodeCollection buildingUseFieldNodeList = buildingUseNode.SelectNodes("./td");
                        Dictionary<string, object> f2vs = new Dictionary<string, object>(); 
                        f2vs.Add("projectId", projectId);
                        f2vs.Add("项目名称", projectName);
                        f2vs.Add("buildingId", buildingId);
                        f2vs.Add("楼名称", buildingName);
                        f2vs.Add("用途", buildingUseFieldNodeList[0].InnerText.Trim());
                        f2vs.Add("批准销售套数", buildingUseFieldNodeList[1].InnerText.Trim());
                        f2vs.Add("批准销售面积", buildingUseFieldNodeList[2].InnerText.Trim());
                        f2vs.Add("已售套数", buildingUseFieldNodeList[3].InnerText.Trim());
                        f2vs.Add("已售面积", buildingUseFieldNodeList[4].InnerText.Trim());
                        f2vs.Add("可售套数", buildingUseFieldNodeList[5].InnerText.Trim());
                        f2vs.Add("可售面积", buildingUseFieldNodeList[6].InnerText.Trim()); 
                        resultEW.AddRow(f2vs);
                    }
                }
            }

            resultEW.SaveToDisk();

            return true;
        }
    }
}