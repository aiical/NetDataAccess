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
    public class GetLoupanPages : ExternalRunWebPage
    {  

        public override bool AfterAllGrab(IListSheet listSheet)
        {
            try
            {
                return this.GetLoupanDetailInfos(listSheet) && this.GetBuildingListPageUrls(listSheet);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        private bool GetBuildingListPageUrls(IListSheet listSheet)
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
            resultColumnDic.Add("pageIndex", 7);
            string resultFilePath = Path.Combine(exportDir, "济南楼盘_楼列表页.xlsx");
            Dictionary<string, string> resultColumnFormat = new Dictionary<string, string>(); 
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, resultColumnFormat);

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;

            Dictionary<string, string> loupanDic = new Dictionary<string, string>();

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
                    HtmlNode pageCountNode = pageHtmlDoc.DocumentNode.SelectSingleNode("//input[@id=\"allpage\"]");
                    if (pageCountNode != null)
                    {
                        int pageCount = int.Parse(pageCountNode.GetAttributeValue("value", ""));

                        for (int j = 0; j < pageCount; j++)
                        {
                            int pageIndex = j + 1;
                            string detailPageUrl = "http://www.jnfdc.gov.cn/onsaling/show_" + pageIndex.ToString() + ".shtml?prjno=" + projectId;
                            Dictionary<string, object> f2vs = new Dictionary<string, object>();
                            f2vs.Add("detailPageUrl", detailPageUrl);
                            f2vs.Add("detailPageName", projectId + "_" + pageIndex.ToString());
                            f2vs.Add("projectId", projectId);
                            f2vs.Add("projectName", projectName);
                            f2vs.Add("pageIndex", pageIndex.ToString());
                            resultEW.AddRow(f2vs);
                        }
                    }
                }
            }

            resultEW.SaveToDisk();

            return true;
        } 

        private bool GetLoupanDetailInfos(IListSheet listSheet)
        { 
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>(); 
            resultColumnDic.Add("项目ID", 0);
            resultColumnDic.Add("项目名称", 1);
            resultColumnDic.Add("项目地址", 2);
            resultColumnDic.Add("企业名称", 3);
            resultColumnDic.Add("所在区县", 4);
            resultColumnDic.Add("项目规模", 5);
            resultColumnDic.Add("总栋数", 6);
            resultColumnDic.Add("可售套数", 7); 
            string resultFilePath = Path.Combine(exportDir, "济南楼盘_楼盘详情.xlsx");
            Dictionary<string, string> resultColumnFormat = new Dictionary<string, string>(); 
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, resultColumnFormat);

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string url = row[detailPageUrlColumnName];
                    string projectId = row["projectId"];
                    string sellable = row["sellable"];

                    HtmlAgilityPack.HtmlDocument pageHtmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                    HtmlNodeCollection trNodeList = pageHtmlDoc.DocumentNode.SelectNodes("//table[@class=\"message_table\"]/tr");


                    string projectName = trNodeList[1].SelectNodes("./td")[1].InnerText.Trim();
                    string address = trNodeList[1].SelectNodes("./td")[3].InnerText.Trim();
                    string companyName = trNodeList[2].SelectNodes("./td")[1].InnerText.Trim();
                    string scope = trNodeList[2].SelectNodes("./td")[3].InnerText.Trim();
                    string projectSize = trNodeList[3].SelectNodes("./td")[1].InnerText.Trim();
                    string buildingCount = trNodeList[3].SelectNodes("./td")[3].InnerText.Trim();

                    Dictionary<string, object> f2vs = new Dictionary<string, object>();
                    f2vs.Add("项目ID", projectId);
                    f2vs.Add("项目名称", projectName);
                    f2vs.Add("项目地址", address);
                    f2vs.Add("企业名称", companyName);
                    f2vs.Add("所在区县", scope);
                    f2vs.Add("项目规模", projectSize);
                    f2vs.Add("总栋数", buildingCount);
                    f2vs.Add("可售套数", sellable);  
                    resultEW.AddRow(f2vs);
                }
            }

            resultEW.SaveToDisk();

            return true;
        } 
    }
}