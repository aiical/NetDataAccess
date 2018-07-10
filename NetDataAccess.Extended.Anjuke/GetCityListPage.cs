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
using HtmlAgilityPack;
using NetDataAccess.Base.Writer;
using NetDataAccess.Base.DB;

namespace NetDataAccess.Extended.Anjuke
{
    public class GetCityListPage : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            try
            {
                this.GetCityList(listSheet);
                this.GetXiaoquFirstPageUrls(listSheet);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return true;
        }

        private void GetXiaoquFirstPageUrls(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("code", 5);
            resultColumnDic.Add("name", 6);
            string resultFilePath = Path.Combine(exportDir, "安居客城市分区.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);

            Dictionary<string, string> urlDic = new Dictionary<string, string>();
             
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"];
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string localFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir);

                    try
                    {
                        HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                        HtmlNodeCollection allCityNodes = htmlDoc.DocumentNode.SelectNodes("//div[@class=\"cl-c-list\"]/ul[@class=\"cl-c-l-ul\"]/li[@class=\"cl-c-l-li\"]/a");

                        for (int j = 0; j < allCityNodes.Count; j++)
                        {
                            HtmlNode cityNode = allCityNodes[j];
                            string url = cityNode.GetAttributeValue("href", "");
                            int cityCodeFromIndex = url.IndexOf("com/") + 4;
                            int cityCodeEndIndex = url.IndexOf("/commu");
                            if (cityCodeEndIndex > 0)
                            {
                                string code = url.Substring(cityCodeFromIndex, cityCodeEndIndex - cityCodeFromIndex);
                                string name = CommonUtil.HtmlDecode(cityNode.InnerText.Trim()).Trim();

                                if (!urlDic.ContainsKey(url))
                                {
                                    urlDic.Add(url, null);
                                    Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                    f2vs.Add("detailPageUrl", url);
                                    f2vs.Add("detailPageName", code);
                                    f2vs.Add("code", code);
                                    f2vs.Add("name", name);
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

        private void GetCityList(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("code", 0);
            resultColumnDic.Add("name", 1);
            resultColumnDic.Add("url", 2);
            string resultFilePath = Path.Combine(exportDir, "安居客城市列表.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);

            Dictionary<string, string> urlDic = new Dictionary<string, string>();
             
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"];
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string localFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir);

                    try
                    {
                        HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                        HtmlNodeCollection allCityNodes = htmlDoc.DocumentNode.SelectNodes("//div[@class=\"cl-c-list\"]/ul[@class=\"cl-c-l-ul\"]/li[@class=\"cl-c-l-li\"]/a");

                        for (int j = 0; j < allCityNodes.Count; j++)
                        {
                            HtmlNode cityNode = allCityNodes[j];
                            string url = cityNode.GetAttributeValue("href", "");
                            int cityCodeFromIndex = url.IndexOf("com/") + 4;
                            int cityCodeEndIndex = url.IndexOf("/commu");
                            if (cityCodeEndIndex > 0)
                            {
                                string code = url.Substring(cityCodeFromIndex, cityCodeEndIndex - cityCodeFromIndex);
                                string name = CommonUtil.HtmlDecode(cityNode.InnerText.Trim()).Trim();
                                if (!urlDic.ContainsKey(url))
                                {
                                    urlDic.Add(url, null);
                                    Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                    f2vs.Add("code", code);
                                    f2vs.Add("name", name);
                                    f2vs.Add("url", url);
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