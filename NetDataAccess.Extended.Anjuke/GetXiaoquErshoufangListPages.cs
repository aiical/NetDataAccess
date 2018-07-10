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
using NetDataAccess.Base.Web;
using System.Net;

namespace NetDataAccess.Extended.Anjuke
{
    public class GetXiaoquErshoufangListPages : ExternalRunWebPage
    {
        public override void WebRequestHtml_BeforeSendRequest(string pageUrl, Dictionary<string, string> listRow, Base.Web.NDAWebClient client)
        {
            string userAgent = this.RunPage.CurrentUserAgents.GetOneUserAgent();
            client.Headers["User-Agent"] = userAgent;
            client.Headers.Add("x-request-with", "XMLHttpRequest");
            client.Headers.Add("cookie", "ctid=" + DateTime.Now.Millisecond);
        }

        public override bool AfterGrabOneCatchException(string pageUrl, Dictionary<string, string> listRow, Exception ex)
        {
            if (ex is EmptyFileException)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public override void CheckRequestCompleteFile(string webPageText, Dictionary<string, string> listRow)
        {
            if (webPageText.Length == 0)
            {

            }
            else if ((webPageText.Contains("<div class=\"fang-item ") || webPageText.Contains("<div class=\"fang-item\">")) && webPageText.Trim().Contains("</div>"))
            {

            }
            else
            {
                throw new Exception("未完全加载文件.");
            }
        }
         

        public override bool AfterAllGrab(IListSheet listSheet)
        {
            try
            {
                this.GetXiaoquInfos(listSheet);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return true;
        }

        private ExcelWriter GetExcelWriter(int fileIndex, string cityName)
        { 
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4); 
            resultColumnDic.Add("xiaoquname", 5);
            resultColumnDic.Add("xiaoquurl", 6);
            resultColumnDic.Add("cityName", 7);
            resultColumnDic.Add("cityCode", 8);
            resultColumnDic.Add("level1AreaName", 9);
            resultColumnDic.Add("level1AreaCode", 10);
            resultColumnDic.Add("level2AreaCode", 11);
            resultColumnDic.Add("level2AreaName", 12);

            string resultFilePath = Path.Combine(exportDir, "安居客在售二手房_" + cityName + "_" + fileIndex.ToString() + ".xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }

        private void GetXiaoquInfos(IListSheet listSheet)
        {
            string[] paramterParts = this.Parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            string cityName = paramterParts[0];

            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            ExcelWriter resultEW = null;
            int fileIndex = 1;

            Dictionary<string, string> fangLinkUrlDic = new Dictionary<string, string>();

            for (int i = 0; i < listSheet.RowCount; i++) 
            {
                if (resultEW == null || resultEW.RowCount > 500000)
                {
                    if (resultEW != null)
                    {
                        resultEW.SaveToDisk();
                    }
                    resultEW = this.GetExcelWriter(fileIndex, cityName);
                    fileIndex++;
                }

                Dictionary<string, string> row = listSheet.GetRow(i); 
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                    try
                    {
                        HtmlNodeCollection fangNodeList = htmlDoc.DocumentNode.SelectNodes("//div[@class=\"baseinfo\"]/a");

                        if (fangNodeList != null)
                        {
                            foreach (HtmlNode fangNode in fangNodeList)
                            {
                                string fangLinkUrl = fangNode.GetAttributeValue("href", "");
                                if (!fangLinkUrlDic.ContainsKey(fangLinkUrl))
                                {
                                    fangLinkUrlDic.Add(fangLinkUrl, null);
                                    Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                    f2vs.Add("detailPageUrl", fangLinkUrl);
                                    f2vs.Add("detailPageName", fangLinkUrl);
                                    f2vs.Add("xiaoquname", row["xiaoquName"]);
                                    f2vs.Add("xiaoquurl", row["xiaoquUrl"]);
                                    f2vs.Add("cityName", row["cityName"]);
                                    f2vs.Add("cityCode", row["cityCode"]);
                                    f2vs.Add("level1AreaName", row["level1AreaName"]);
                                    f2vs.Add("level1AreaCode", row["level1AreaCode"]);
                                    f2vs.Add("level2AreaCode", row["level2AreaCode"]);
                                    f2vs.Add("level2AreaName", row["level2AreaName"]);

                                    resultEW.AddRow(f2vs);
                                }
                            }
                        } 
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
            }
            resultEW.SaveToDisk();
        }
    }
}