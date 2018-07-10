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

namespace NetDataAccess.Extended.Yiguo
{
    public class LiangxianDetailPageUrl : CustomProgramBase
    {
        public bool Run(string parameters, IListSheet listSheet)
        {
            return GetAllDetailPageUrl(listSheet);
        }
        private bool GetAllDetailPageUrl(IListSheet listSheet)
        {
            string exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("productSysNo", 5);
            resultColumnDic.Add("category1Code", 6);
            resultColumnDic.Add("category2Code", 7);
            resultColumnDic.Add("category3Code", 8);
            resultColumnDic.Add("category1Name", 9);
            resultColumnDic.Add("category2Name", 10);
            resultColumnDic.Add("category3Name", 11);
            resultColumnDic.Add("district", 12);
            string resultFilePath = Path.Combine(exportDir, "两鲜网获取所有详情页.xlsx");
            
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic);
             

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;
            string categoryNameColumnName = SysConfig.DetailPageNameFieldName;

            Dictionary<string, Dictionary<string, Dictionary<string, string>>> categoryFullNameToDetailUrls = new Dictionary<string, Dictionary<string, Dictionary<string, string>>>();

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string url = row[detailPageUrlColumnName];
                    string categoryFullName = row[categoryNameColumnName];
                    string category1Code = row["category1Code"];
                    string category2Code = row["category2Code"];
                    string category3Code = row["category3Code"];
                    string category1Name = row["category1Name"];
                    string category2Name = row["category2Name"];
                    string category3Name = row["category3Name"]; 
                    string district = row["district"];
                    string cookie = row["cookie"]; 
                    string localFilePath = this.RunPage.GetFilePath(url, pageSourceDir);
                    TextReader tr = null;

                    Dictionary<string, Dictionary<string, string>> detailUrlToProperties = new Dictionary<string, Dictionary<string, string>>();
                    categoryFullNameToDetailUrls.Add(categoryFullName, detailUrlToProperties);

                    Dictionary<string, Dictionary<string, string>> c1FullNameToProperties = null;
                    Dictionary<string, Dictionary<string, string>> c2FullNameToProperties = null;
                    if (!CommonUtil.IsNullOrBlank(category2Name) && categoryFullNameToDetailUrls.ContainsKey(category1Name))
                    {
                        c1FullNameToProperties = categoryFullNameToDetailUrls[category1Name];
                    }
                    if (!CommonUtil.IsNullOrBlank(category3Name) && categoryFullNameToDetailUrls.ContainsKey(category1Name + "_" + category2Name))
                    { 
                        c2FullNameToProperties = categoryFullNameToDetailUrls[category1Name + "_" + category2Name];
                    }

                    try
                    {
                        tr = new StreamReader(localFilePath);
                        string webPageHtml = tr.ReadToEnd();

                        HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                        htmlDoc.LoadHtml(webPageHtml);

                        HtmlNodeCollection allDetailPageNodesA = htmlDoc.DocumentNode.SelectNodes("//section[@class=\"category-products\"]/ul/li/div/div/div[3]/div[@class=\"left\"]/h2/a");

                        HtmlNodeCollection allDetailPageNodesB = htmlDoc.DocumentNode.SelectNodes("//section[@class=\"category-products\"]/ul/li/div/div/div[2]/h2/a");

                        List<HtmlNode> allDetailPageNodes = new List<HtmlNode>();
                        if (allDetailPageNodesA != null)
                        {
                            allDetailPageNodes.AddRange(allDetailPageNodesA);
                        }
                        if (allDetailPageNodesB != null)
                        {
                            allDetailPageNodes.AddRange(allDetailPageNodesB);
                        }

                        foreach (HtmlNode detailPageNode in allDetailPageNodes)
                        {
                            string goodDetailPageUrl = detailPageNode.Attributes["href"].Value;
                            string detailPageUrl = goodDetailPageUrl;
                            string detailPageName = detailPageUrl;
                            Dictionary<string, string> f2vs = new Dictionary<string, string>();
                            f2vs.Add("detailPageUrl", detailPageUrl);
                            f2vs.Add("detailPageName", detailPageName);
                            f2vs.Add("cookie", cookie);
                            f2vs.Add("category1Code", category1Code);
                            f2vs.Add("category2Code", category2Code);
                            f2vs.Add("category3Code", category3Code);
                            f2vs.Add("category1Name", category1Name);
                            f2vs.Add("category2Name", category2Name);
                            f2vs.Add("category3Name", category3Name);
                            f2vs.Add("district", district);
                            f2vs.Add("productSysNo", detailPageName);
                            detailUrlToProperties.Add(detailPageUrl, f2vs);

                            if (c1FullNameToProperties != null && c1FullNameToProperties.ContainsKey(detailPageUrl))
                            {
                                c1FullNameToProperties.Remove(detailPageUrl);
                            }
                            if (c2FullNameToProperties != null && c2FullNameToProperties.ContainsKey(detailPageUrl))
                            {
                                c2FullNameToProperties.Remove(detailPageUrl);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        if (tr != null)
                        {
                            tr.Dispose();
                            tr = null;
                        }
                        this.RunPage.InvokeAppendLogText("读取出错.  " + ex.Message + " LocalPath = " + localFilePath, LogLevelType.Error, true);
                        throw ex;
                    }
                }
            }

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string url = row[detailPageUrlColumnName];
                    string categoryFullName = row[categoryNameColumnName];
                    Dictionary<string, Dictionary<string, string>> detailUrlToProperties = categoryFullNameToDetailUrls[categoryFullName];
                    foreach (string detailUrl in detailUrlToProperties.Keys)
                    {
                        Dictionary<string, string> f2vs = detailUrlToProperties[detailUrl];
                        resultEW.AddRow(f2vs);
                    }
                }
            }
            

            resultEW.SaveToDisk(); 
            
            return true;
        }
    }
}