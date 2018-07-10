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
    public class GetCities : ExternalRunWebPage
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
            resultColumnDic.Add("cityCode", 5);
            resultColumnDic.Add("city", 6);
            resultColumnDic.Add("parentDistrict", 7);
            resultColumnDic.Add("baseUrl", 8);
            string resultFilePath = Path.Combine(exportDir, "大众点评城市首页.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string typeCode = row["typeCode"];
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    //特区直辖市
                    HtmlAgilityPack.HtmlDocument pageHtmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                    HtmlNodeCollection termCityNodeList = pageHtmlDoc.DocumentNode.SelectNodes("//div[@class=\"terms\"]");
                    for (int j = 0; j < termCityNodeList.Count; j++)
                    {
                        HtmlNode termNode = termCityNodeList[j];  
                        HtmlNodeCollection cityNodeList = termNode.SelectNodes("./a");
                        for (int k = 0; k < cityNodeList.Count; k++)
                        {
                            HtmlNode cityNode = cityNodeList[k];
                            string cityName = cityNode.InnerText.Trim();
                            string cityUrl = cityNode.GetAttributeValue("href", "");
                            int codeBeginIndex = cityUrl.LastIndexOf("/") + 1;
                            string cityCode = cityUrl.Substring(codeBeginIndex);

                            string pageUrl = "http://www.dianping.com/" + cityCode + "/" + typeCode;

                            Dictionary<string, string> f2vs = new Dictionary<string, string>();
                            f2vs.Add("detailPageUrl", pageUrl);
                            f2vs.Add("detailPageName", cityCode + "_" + typeCode);
                            f2vs.Add("cityCode", cityCode);
                            f2vs.Add("city", cityName);
                            f2vs.Add("parentDistrict", cityName);
                            f2vs.Add("baseUrl", pageUrl);
                            resultEW.AddRow(f2vs);
                        }
                    }

                    //省自治区
                    HtmlNodeCollection termNodeList = pageHtmlDoc.DocumentNode.SelectNodes("//dl[@class=\"terms\"]");
                    for (int j = 0; j < termNodeList.Count; j++)
                    {
                        HtmlNode termNode = termNodeList[j];
                        HtmlNode parentDistrictNode = termNode.SelectSingleNode("./dt");
                        string parentDistrict = parentDistrictNode.InnerText.Trim();

                        HtmlNodeCollection cityNodeList = termNode.SelectNodes("./dd/a");
                        for (int k = 0; k < cityNodeList.Count; k++)
                        {
                            HtmlNode cityNode = cityNodeList[k];
                            string cityName = cityNode.InnerText.Trim();
                            string cityUrl = cityNode.GetAttributeValue("href", "");
                            int codeBeginIndex = cityUrl.LastIndexOf("/") + 1;
                            string cityCode = cityUrl.Substring(codeBeginIndex);

                            string pageUrl = "http://www.dianping.com/" + cityCode + "/" + typeCode;

                            Dictionary<string, string> f2vs = new Dictionary<string, string>();
                            f2vs.Add("detailPageUrl", pageUrl);
                            f2vs.Add("detailPageName", cityCode + "_" + typeCode);
                            f2vs.Add("cityCode", cityCode);
                            f2vs.Add("city", cityName);
                            f2vs.Add("parentDistrict", parentDistrict);
                            f2vs.Add("baseUrl", pageUrl);
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