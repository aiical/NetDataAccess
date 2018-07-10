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
using System.Net;

namespace NetDataAccess.Extended.Dzdp
{
    public class GetCategoryAndDistrict : ExternalRunWebPage
    {
        public override void CheckRequestCompleteFile(string webPageText, Dictionary<string, string> listRow)
        {
            string city = listRow["city"];
            if (!webPageText.Contains("'"+city+"'"))
            {
                throw new GiveUpException("无法获取到对应页面，放弃");
            }
        }

        public override bool AfterGrabOneCatchException(string pageUrl, Dictionary<string, string> listRow, Exception ex)
        {
            if (ex.InnerException is GiveUpException)
            {
                this.RunPage.InvokeAppendLogText(ex.Message + ", pageUrl = " + pageUrl, LogLevelType.Error, true);
                return true;
            }
            else if (ex.InnerException is WebException)
            {
                WebException webEx = (WebException)ex.InnerException;
                if (webEx.Status == WebExceptionStatus.ProtocolError)
                {
                    HttpWebResponse webRes = (HttpWebResponse)webEx.Response;
                    if (webRes.StatusCode == HttpStatusCode.NotFound)
                    {
                        this.RunPage.InvokeAppendLogText("服务器端不存在此网页(404), pageUrl = " + pageUrl, LogLevelType.Error, true);
                        return true;
                    }
                }
            }
            return false;
        }
        public override bool AfterAllGrab(IListSheet listSheet)
        { 
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("infoValue", 5);
            resultColumnDic.Add("infoName", 6);
            resultColumnDic.Add("infoType", 7);
            resultColumnDic.Add("city", 8);
            resultColumnDic.Add("baseUrl", 9);
            string resultFilePath = Path.Combine(exportDir, "大众点评获取二级分类.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string url = row[detailPageUrlColumnName];
                    string cookie = row["cookie"];
                    string city = row["city"];
                    string baseUrl = row["baseUrl"];
                    HtmlAgilityPack.HtmlDocument pageHtmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                    try
                    {
                        HtmlNodeCollection allGNodes = pageHtmlDoc.DocumentNode.SelectNodes("//div[@id=\"classfy\"]/a");
                        HtmlNodeCollection allRNodes = pageHtmlDoc.DocumentNode.SelectNodes("//div[@id=\"region-nav\"]/a");
                        if (allGNodes != null && allRNodes != null)
                        {
                            foreach (HtmlNode gNode in allGNodes)
                            {
                                string nodeHref = gNode.Attributes["href"].Value;
                                string infoName = gNode.InnerText;
                                SaveRow("G", infoName, cookie, baseUrl, nodeHref, city, resultEW);
                            }
                            foreach (HtmlNode rNode in allRNodes)
                            {
                                string nodeHref = rNode.Attributes["href"].Value;
                                string infoName = rNode.InnerText;
                                SaveRow("R", infoName, cookie, baseUrl, nodeHref, city, resultEW);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        string dir = this.RunPage.GetDetailSourceFileDir();
                        string fileUrl = this.RunPage.GetFilePath(url, dir);
                        File.Delete(fileUrl);
                        this.RunPage.InvokeAppendLogText("删除文件: " + fileUrl, LogLevelType.Error, true);
                        //throw ex;
                    }
                }
            }
            resultEW.SaveToDisk();

            return true;
        }
        private void SaveRow(string infoType, string infoName, string cookie, string baseUrl, string nodeHref, string city, ExcelWriter resultEW)
        {
            string[] infoPieces = nodeHref.Split(new string[] { "/" }, StringSplitOptions.RemoveEmptyEntries);
            string infoValue = infoPieces[infoPieces.Length - 1];
            string detailPageName = city + "_" + infoValue;
            string detailPageUrl = baseUrl + "/" + infoValue;
            Dictionary<string, string> f2vs = new Dictionary<string, string>();
            f2vs.Add("detailPageUrl", detailPageUrl);
            f2vs.Add("detailPageName", detailPageName);
            f2vs.Add("cookie", cookie);
            f2vs.Add("infoValue", infoValue);
            f2vs.Add("infoType", infoType);
            f2vs.Add("infoName", infoName);
            f2vs.Add("city", city);
            f2vs.Add("baseUrl", baseUrl);
            resultEW.AddRow(f2vs);
        }
    }
}