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

namespace NetDataAccess.Extended.Dzdp
{
    public class GetSubCategoryAndSubDistrict : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        { 
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("city", 5);
            resultColumnDic.Add("g", 6);
            resultColumnDic.Add("r", 7);
            resultColumnDic.Add("gName", 8);
            resultColumnDic.Add("rName", 9);
            string resultFilePath = Path.Combine(exportDir, "大众点评获取列表页首页.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;

            Dictionary<string, Dictionary<string, string>> cityToGs = new Dictionary<string, Dictionary<string, string>>();
            Dictionary<string, Dictionary<string, string>> cityToRs = new Dictionary<string, Dictionary<string, string>>();
            Dictionary<string, string> cityToBaseUrls = new Dictionary<string, string>();
            Dictionary<string, string> cityToCookies = new Dictionary<string, string>();


            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string url = row[detailPageUrlColumnName];
                    string cookie = row["cookie"];
                    string city = row["city"];
                    string infoValue = row["infoValue"];
                    string infoName = row["infoName"];
                    string baseUrl = row["baseUrl"];
                    if (!cityToGs.ContainsKey(city))
                    {
                        cityToGs.Add(city, new Dictionary<string, string>());
                        cityToRs.Add(city, new Dictionary<string, string>());
                        cityToBaseUrls.Add(city, baseUrl);
                        cityToCookies.Add(city, cookie);
                    }

                    Dictionary<string, string> g2Names = cityToGs[city];
                    Dictionary<string, string> r2Names = cityToRs[city];

                    HtmlAgilityPack.HtmlDocument pageHtmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                    HtmlNodeCollection allSubGNodes = pageHtmlDoc.DocumentNode.SelectNodes("//div[@id=\"classfy-sub\"]/a");
                    HtmlNodeCollection allSubRNodes = pageHtmlDoc.DocumentNode.SelectNodes("//div[@id=\"region-nav-sub\"]/a");

                    if (allSubGNodes != null)
                    {
                        foreach (HtmlNode subGNode in allSubGNodes)
                        {
                            string subNodeHref = subGNode.Attributes["href"].Value;
                            string subInfoName = subGNode.InnerText;
                            if (subInfoName == "不限")
                            {
                                subInfoName = infoName;
                            }
                            SaveRow(subInfoName, subNodeHref, g2Names);
                        }
                    }
                    if (allSubRNodes != null)
                    {
                        foreach (HtmlNode subRNode in allSubRNodes)
                        {
                            string subNodeHref = subRNode.Attributes["href"].Value;
                            string subInfoName = subRNode.InnerText;
                            if (subInfoName == "不限")
                            {
                                subInfoName = infoName;
                            }
                            SaveRow(subInfoName, subNodeHref, r2Names);
                        }
                    }
                }
            }
            foreach (string city in cityToGs.Keys)
            {
                Dictionary<string, string> g2Names = cityToGs[city];
                Dictionary<string, string> r2Names = cityToRs[city];
                string baseUrl = cityToBaseUrls[city];
                string cookie = cityToCookies[city];
                foreach (string g in g2Names.Keys)
                {
                    string gName = g2Names[g];
                    foreach (string r in r2Names.Keys)
                    {
                        string rName = r2Names[r];
                        string subUrl = g + r;
                        string detailPageUrl = baseUrl + "/" + subUrl;
                        Dictionary<string, string> f2vs = new Dictionary<string, string>();
                        f2vs.Add("detailPageUrl", detailPageUrl);
                        f2vs.Add("detailPageName", city + "_" + subUrl);
                        f2vs.Add("cookie", cookie);
                        f2vs.Add("city", city);
                        f2vs.Add("g", g);
                        f2vs.Add("r", r);
                        f2vs.Add("gName", gName);
                        f2vs.Add("rName", rName);
                        resultEW.AddRow(f2vs);
                    }
                }
            }


            resultEW.SaveToDisk();

            return true;
        }
        private void SaveRow(string infoName, string nodeHref, Dictionary<string, string> code2Names)
        {
            string[] infoPieces = nodeHref.Split(new string[] { "/" }, StringSplitOptions.RemoveEmptyEntries);
            string infoValue = infoPieces[infoPieces.Length - 1];
            if (!code2Names.ContainsKey(infoValue))
            {
                code2Names.Add(infoValue, infoName);
            }
        }
    }
}