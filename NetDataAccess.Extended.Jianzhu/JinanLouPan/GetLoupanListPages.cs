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
    public class GetLoupanListPages : ExternalRunWebPage
    {  

        public override bool AfterAllGrab(IListSheet listSheet)
        {
            try
            {
                return this.GetLoupanInfos(listSheet);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private bool GetLoupanInfos(IListSheet listSheet)
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
            resultColumnDic.Add("address", 7);
            resultColumnDic.Add("companyName", 8);
            resultColumnDic.Add("sellable", 9);
            string resultFilePath = Path.Combine(exportDir, "济南楼盘_楼盘详情页.xlsx");
            Dictionary<string, string> resultColumnFormat = new Dictionary<string, string>();
            resultColumnFormat.Add("sellable", "#,##0");
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

                    HtmlAgilityPack.HtmlDocument pageHtmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                    HtmlNodeCollection itemNodeList = pageHtmlDoc.DocumentNode.SelectNodes("//table[@class=\"project_table\"]/tr");
                    for (int j = 1; j < itemNodeList.Count - 1; j++)
                    {
                        HtmlNode itemNode = itemNodeList[j];
                        HtmlNodeCollection fieldNodeList = itemNode.SelectNodes("./td");
                        HtmlNode projectNameNode = fieldNodeList[1];
                        string projectName = projectNameNode.GetAttributeValue("title", "");
                        string loupanPartUrl = projectNameNode.SelectSingleNode("./a").GetAttributeValue("href", "");
                        int equalIndex = loupanPartUrl.LastIndexOf("=");
                        string projectId = loupanPartUrl.Substring(equalIndex + 1).Trim();

                        string address = fieldNodeList[2].InnerText.Trim();
                        string companyName = fieldNodeList[3].InnerText.Trim();
                        int sellable = int.Parse(fieldNodeList[4].InnerText.Trim());
                        if (!loupanDic.ContainsKey(projectId))
                        {
                            loupanDic.Add(projectId, null);
                            string detailPageUrl = "http://www.jnfdc.gov.cn" + loupanPartUrl;
                            Dictionary<string, object> f2vs = new Dictionary<string, object>();
                            f2vs.Add("detailPageUrl", detailPageUrl);
                            f2vs.Add("detailPageName", projectId);
                            f2vs.Add("projectId", projectId);
                            f2vs.Add("projectName", projectName);
                            f2vs.Add("address", address);
                            f2vs.Add("companyName", companyName);
                            f2vs.Add("sellable", sellable);
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