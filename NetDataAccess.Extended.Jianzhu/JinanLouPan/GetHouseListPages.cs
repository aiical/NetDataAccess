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
    public class GetHouseListPages : ExternalRunWebPage
    {
        public override void CheckRequestCompleteFile(string webPageText, Dictionary<string, string> listRow)
        {
            HtmlAgilityPack.HtmlDocument pageHtmlDoc = new HtmlAgilityPack.HtmlDocument();
            pageHtmlDoc.LoadHtml(webPageText);
            HtmlNodeCollection floorTrNodeList = pageHtmlDoc.DocumentNode.SelectNodes("//table[@id=\"floorTable\"]/tbody/tr"); 
            HtmlNodeCollection bussTrNodeList = pageHtmlDoc.DocumentNode.SelectNodes("//table[@id=\"bussTable\"]/tbody/tr");
            if ((floorTrNodeList == null || floorTrNodeList.Count == 0) && (bussTrNodeList == null || bussTrNodeList.Count == 0))
            {
                throw new Exception("未能完整获取页面");
            } 
        }

        public override bool AfterAllGrab(IListSheet listSheet)
        {
            try
            {
                return this.GetHouseListInfos(listSheet) ;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        private ExcelWriter CreateResultExcelWriter(int fileIndex)
        {
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("projectId", 5);
            resultColumnDic.Add("项目名称", 6);
            resultColumnDic.Add("buildingId", 7);
            resultColumnDic.Add("楼名称", 8);
            resultColumnDic.Add("是否住宅房屋", 9);
            resultColumnDic.Add("单元号", 10);
            resultColumnDic.Add("顺序号", 11);
            resultColumnDic.Add("楼层", 12);
            resultColumnDic.Add("houseId", 13);
            resultColumnDic.Add("houseName", 14);
            string resultFilePath = Path.Combine(exportDir, "济南楼盘_房间信息_" + fileIndex + ".xlsx");
            Dictionary<string, string> resultColumnFormat = new Dictionary<string, string>();
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, resultColumnFormat);
            return resultEW;
        }

        private bool GetHouseListInfos(IListSheet listSheet)
        {

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;

            Dictionary<string, string> houseDic = new Dictionary<string, string>();

            int fileIndex = 1;
            ExcelWriter resultEW = this.CreateResultExcelWriter(fileIndex);

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                if (resultEW.RowCount >= 500000)
                {
                    resultEW.SaveToDisk();
                    fileIndex++;
                    resultEW = this.CreateResultExcelWriter(fileIndex);
                }

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
                    HtmlNodeCollection floorTrNodeList = pageHtmlDoc.DocumentNode.SelectNodes("//table[@id=\"floorTable\"]/tbody/tr");
                    if (floorTrNodeList != null && floorTrNodeList.Count > 2)
                    {
                        //包含了住宅房屋

                        List<string> danYuanList = new List<string>();
                        HtmlNodeCollection danYuanTdNodeList = floorTrNodeList[0].SelectNodes("./td");
                        for (int j = 1; j < danYuanTdNodeList.Count; j++)
                        {
                            HtmlNode danYuanTdNode = danYuanTdNodeList[j];
                            int colspan = int.Parse(danYuanTdNode.GetAttributeValue("colspan", "1"));
                            string danYuanName = danYuanTdNode.InnerText.Trim();
                            for (int k = 0; k < colspan; k++)
                            {
                                danYuanList.Add(danYuanName);
                            }
                        }

                        List<string> shunXuHaoList = new List<string>();
                        HtmlNodeCollection shunXuHaoTdNodeList = floorTrNodeList[1].SelectNodes("./td");
                        for (int j = 1; j < shunXuHaoTdNodeList.Count; j++)
                        {
                            HtmlNode shunXuHaoTdNode = shunXuHaoTdNodeList[j];
                            string shunXuHao = shunXuHaoTdNode.InnerText.Trim();
                            shunXuHaoList.Add(shunXuHao);
                        }

                        for (int j = 2; j < floorTrNodeList.Count; j++)
                        {
                            HtmlNodeCollection houseTdNodeList = floorTrNodeList[j].SelectNodes("./td");
                            string floorName = houseTdNodeList[0].InnerText.Trim();
                            for (int k = 1; k < houseTdNodeList.Count; k++)
                            {
                                HtmlNode houseTdNode = houseTdNodeList[k];
                                string houseId = houseTdNode.GetAttributeValue("id", "").Trim();
                                string houseName = houseTdNode.InnerText.Trim();

                                string danYuanName = danYuanList[k - 1];
                                string shunXuHao = shunXuHaoList[k - 1];

                                string detailPageUrl = "http://www.jnfdc.gov.cn/onsaling/viewDiv.shtml?fid=" + houseId;

                                if (houseId.Length > 0 && !houseDic.ContainsKey(houseId))
                                {
                                    houseDic.Add(houseId, null);
                                    Dictionary<string, object> f2vs = new Dictionary<string, object>();
                                    f2vs.Add("detailPageUrl", detailPageUrl);
                                    f2vs.Add("detailPageName", houseId);
                                    f2vs.Add("projectId", projectId);
                                    f2vs.Add("项目名称", projectName);
                                    f2vs.Add("buildingId", buildingId);
                                    f2vs.Add("楼名称", buildingName);
                                    f2vs.Add("是否住宅房屋", "是");
                                    f2vs.Add("单元号", danYuanName);
                                    f2vs.Add("顺序号", shunXuHao);
                                    f2vs.Add("楼层", floorName);
                                    f2vs.Add("houseId", houseId);
                                    f2vs.Add("houseName", houseName);
                                    resultEW.AddRow(f2vs);
                                }
                            }
                        }
                    }

                    HtmlNodeCollection bussTrNodeList = pageHtmlDoc.DocumentNode.SelectNodes("//table[@id=\"bussTable\"]/tbody/tr");
                    if (bussTrNodeList != null && bussTrNodeList.Count > 2)
                    {
                        //包含了非住宅房屋 

                        string floorName = "";
                        int floorIndex = 1;

                        for (int j = 0; j < bussTrNodeList.Count; j++)
                        {
                            HtmlNodeCollection houseTdNodeList = bussTrNodeList[j].SelectNodes("./td");
                            string tempName = houseTdNodeList[0].InnerText.Trim();
                            if (tempName.Length > 0)
                            {
                                floorName = tempName + "_" + floorIndex.ToString();
                                floorIndex++;
                            }
                            for (int k = 1; k < houseTdNodeList.Count; k++)
                            {
                                HtmlNode houseTdNode = houseTdNodeList[k];
                                string houseId = houseTdNode.GetAttributeValue("id", "").Trim();
                                string houseName = houseTdNode.InnerText.Trim();

                                string detailPageUrl = "http://www.jnfdc.gov.cn/onsaling/viewDiv.shtml?fid=" + houseId;
                                if (houseId.Length > 0 && !houseDic.ContainsKey(houseId))
                                {
                                    houseDic.Add(houseId, null);
                                    Dictionary<string, object> f2vs = new Dictionary<string, object>();
                                    f2vs.Add("detailPageUrl", detailPageUrl);
                                    f2vs.Add("detailPageName", houseId);
                                    f2vs.Add("projectId", projectId);
                                    f2vs.Add("项目名称", projectName);
                                    f2vs.Add("buildingId", buildingId);
                                    f2vs.Add("楼名称", buildingName);
                                    f2vs.Add("是否住宅房屋", "否");
                                    f2vs.Add("楼层", floorName);
                                    f2vs.Add("houseId", houseId);
                                    f2vs.Add("houseName", houseName);
                                    resultEW.AddRow(f2vs);
                                }
                            }
                        }
                    }

                }
            }

            resultEW.SaveToDisk();

            return true;
        }
         
    }
}