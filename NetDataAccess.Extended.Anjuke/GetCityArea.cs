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

namespace NetDataAccess.Extended.Anjuke
{
    public class GetCityArea : ExternalRunWebPage
    {
        public override void WebRequestHtml_BeforeSendRequest(string pageUrl, Dictionary<string, string> listRow, Base.Web.NDAWebClient client)
        {
            string userAgent = this.RunPage.CurrentUserAgents.GetOneUserAgent();
            client.Headers["User-Agent"] = userAgent;
        }
        public override void CheckRequestCompleteFile(string webPageText, Dictionary<string, string> listRow)
        {
            string name = listRow["name"];
            if (webPageText.Contains(name) && !webPageText.Contains("访问验证") && webPageText.Trim().EndsWith("</html>"))
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
                this.GetCityAreaLevel1List(listSheet);
                this.GetCityAreaLevel2PageUrls(listSheet);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return true;
        }

        private void GetCityAreaLevel2PageUrls(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("cityCode", 5);
            resultColumnDic.Add("cityName", 6);
            resultColumnDic.Add("level1AreaCode", 7);
            resultColumnDic.Add("level1AreaName", 8);
            resultColumnDic.Add("level2AreaCode", 9);
            resultColumnDic.Add("level2AreaName", 10);
            string resultFilePath = Path.Combine(exportDir, "安居客城市分区小区列表页.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null); 

            Dictionary<string, string> urlDic = new Dictionary<string, string>();

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"];
                string cityName = row["name"];
                string cityCode = row["code"];
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string localFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir);

                    HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                    
                    HtmlNodeCollection allAreaNodes = htmlDoc.DocumentNode.SelectNodes("//ul[@class=\"regionlist\"]/li");
                    try
                    {
                        for (int j = 0; j < allAreaNodes.Count; j++)
                        {
                            HtmlNode areaNode = allAreaNodes[j];
                            if (j == 0)
                            {
                                if (allAreaNodes.Count == 1)
                                {

                                    string url = areaNode.GetAttributeValue("data-href", "");
                                    if (!urlDic.ContainsKey(url))
                                    {
                                        urlDic.Add(url, null);
                                        Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                        f2vs.Add("detailPageUrl", url);
                                        f2vs.Add("detailPageName", cityCode + "_" + cityCode + "_" + cityCode);
                                        f2vs.Add("cityCode", cityCode);
                                        f2vs.Add("cityName", cityName);
                                        f2vs.Add("level1AreaCode", cityCode);
                                        f2vs.Add("level1AreaName", cityName);
                                        f2vs.Add("level2AreaCode", cityCode);
                                        f2vs.Add("level2AreaName", cityName);
                                        resultEW.AddRow(f2vs);
                                    }
                                }
                            }
                            else
                            {
                                string dataId = areaNode.GetAttributeValue("data-id", "");
                                string level1AreaName = CommonUtil.HtmlDecode(areaNode.InnerText.Trim()).Trim();

                                HtmlNodeCollection allSubAreaNodes = htmlDoc.DocumentNode.SelectNodes("//div[@id=\"blockinfo-" + dataId + "\"]/div/a");

                                if (allSubAreaNodes != null)
                                {

                                    HtmlNode subAreaQuanbuNode = allSubAreaNodes[0];
                                    string url = subAreaQuanbuNode.GetAttributeValue("href", "");
                                    int areaCodeEndIndex = url.LastIndexOf("/");
                                    int areaCodeFromIndex = url.LastIndexOf("/", areaCodeEndIndex - 1) + 1;
                                    string level1AreaCode = url.Substring(areaCodeFromIndex, areaCodeEndIndex - areaCodeFromIndex);

                                    for (int k = 0; k < allSubAreaNodes.Count; k++)
                                    {
                                        if (allSubAreaNodes.Count > 1 && k == 0)
                                        {
                                            //忽略掉“全部”节点
                                            continue;
                                        }
                                        else
                                        {
                                            HtmlNode subAreaNode = allSubAreaNodes[k];
                                            string level2Url = subAreaNode.GetAttributeValue("href", "");
                                            int level2AreaCodeEndIndex = level2Url.LastIndexOf("/");
                                            int level2AreaCodeFromIndex = level2Url.LastIndexOf("/", level2AreaCodeEndIndex - 1) + 1;
                                            string level2AreaCode = level2Url.Substring(level2AreaCodeFromIndex, level2AreaCodeEndIndex - level2AreaCodeFromIndex);
                                            string level2AreaName = CommonUtil.HtmlDecode(subAreaNode.InnerText.Trim()).Trim();

                                            if (!urlDic.ContainsKey(level2Url))
                                            {
                                                urlDic.Add(level2Url, null);
                                                Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                                f2vs.Add("detailPageUrl", level2Url);
                                                f2vs.Add("detailPageName", cityCode + "_" + level1AreaCode + "_" + level2AreaCode);
                                                f2vs.Add("cityCode", cityCode);
                                                f2vs.Add("cityName", cityName);
                                                f2vs.Add("level1AreaCode", level1AreaCode);
                                                f2vs.Add("level1AreaName", level1AreaName);
                                                f2vs.Add("level2AreaCode", level2AreaCode);
                                                f2vs.Add("level2AreaName", level2AreaName);
                                                resultEW.AddRow(f2vs);
                                            }
                                        }
                                    }
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

        private void GetCityAreaLevel1List(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("cityCode", 0);
            resultColumnDic.Add("cityName", 1);
            resultColumnDic.Add("level1AreaCode", 2);
            resultColumnDic.Add("level1AreaName", 3);
            resultColumnDic.Add("level2AreaCode", 4);
            resultColumnDic.Add("level2AreaName", 5); 
            resultColumnDic.Add("url", 6);
            string resultFilePath = Path.Combine(exportDir, "安居客城市分区列表.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);

            Dictionary<string, string> urlDic = new Dictionary<string, string>();
             
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"];
                string cityName = row["name"];
                string cityCode = row["code"];
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string localFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir);

                    HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);


                    HtmlNodeCollection allAreaNodes = htmlDoc.DocumentNode.SelectNodes("//ul[@class=\"regionlist\"]/li");
                    try
                    {
                        for (int j = 0; j < allAreaNodes.Count; j++)
                        {
                            HtmlNode areaNode = allAreaNodes[j];
                            if (j == 0)
                            {
                                if (allAreaNodes.Count == 1)
                                {

                                    string url = areaNode.GetAttributeValue("data-href", "");
                                    if (!urlDic.ContainsKey(url))
                                    {
                                        urlDic.Add(url, null);
                                        Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                        f2vs.Add("cityCode", cityCode);
                                        f2vs.Add("cityName", cityName);
                                        f2vs.Add("level1AreaCode", cityCode);
                                        f2vs.Add("level1AreaName", cityName);
                                        f2vs.Add("level2AreaCode", cityCode);
                                        f2vs.Add("level2AreaName", cityName);
                                        f2vs.Add("url", url);
                                        resultEW.AddRow(f2vs);
                                    }
                                }
                            }
                            else
                            {
                                string dataId = areaNode.GetAttributeValue("data-id", "");
                                string level1AreaName = CommonUtil.HtmlDecode(areaNode.InnerText.Trim()).Trim();

                                HtmlNodeCollection allSubAreaNodes = htmlDoc.DocumentNode.SelectNodes("//div[@id=\"blockinfo-" + dataId + "\"]/div/a");

                                if (allSubAreaNodes != null)
                                {
                                    try
                                    {
                                        HtmlNode subAreaQuanbuNode = allSubAreaNodes[0];
                                        string url = subAreaQuanbuNode.GetAttributeValue("href", "");
                                        int areaCodeEndIndex = url.LastIndexOf("/");
                                        int areaCodeFromIndex = url.LastIndexOf("/", areaCodeEndIndex - 1) + 1;
                                        string level1AreaCode = url.Substring(areaCodeFromIndex, areaCodeEndIndex - areaCodeFromIndex);

                                        for (int k = 0; k < allSubAreaNodes.Count; k++)
                                        {
                                            if (allSubAreaNodes.Count > 1 && k == 0)
                                            {
                                                //忽略掉“全部”节点
                                                continue;
                                            }
                                            else
                                            {
                                                HtmlNode subAreaNode = allSubAreaNodes[k];
                                                string level2Url = subAreaNode.GetAttributeValue("href", "");
                                                try
                                                {
                                                    int level2AreaCodeEndIndex = level2Url.LastIndexOf("/");
                                                    int level2AreaCodeFromIndex = level2Url.LastIndexOf("/", level2AreaCodeEndIndex - 1) + 1;
                                                    string level2AreaCode = level2Url.Substring(level2AreaCodeFromIndex, level2AreaCodeEndIndex - level2AreaCodeFromIndex);
                                                    string level2AreaName = CommonUtil.HtmlDecode(subAreaNode.InnerText.Trim()).Trim();

                                                    if (!urlDic.ContainsKey(level2Url))
                                                    {
                                                        urlDic.Add(level2Url, null);
                                                        Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                                        f2vs.Add("cityCode", cityCode);
                                                        f2vs.Add("cityName", cityName);
                                                        f2vs.Add("level1AreaCode", level1AreaCode);
                                                        f2vs.Add("level1AreaName", level1AreaName);
                                                        f2vs.Add("level2AreaCode", level2AreaCode);
                                                        f2vs.Add("level2AreaName", level2AreaName);
                                                        f2vs.Add("url", level2Url);
                                                        resultEW.AddRow(f2vs);
                                                    }
                                                }
                                                catch (Exception ex)
                                                {
                                                    throw ex;
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