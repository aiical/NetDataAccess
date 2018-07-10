using HtmlAgilityPack;
using NetDataAccess.Base.Common;
using NetDataAccess.Base.Config;
using NetDataAccess.Base.DB;
using NetDataAccess.Base.UI;
using NetDataAccess.Base.Writer;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace NetDataAccess.Extended.Xingzheng.TongJiYongQuHuaDaiMa
{
    public class Sheng : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            this.GetProvinces(listSheet);
            return true;
        }

        private void GetProvinces(IListSheet listSheet)
        {
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            ExcelWriter ew = this.GetExcelWriter();

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string detailUrl = row["detailPageUrl"];
                    string year = row["year"];
                    try
                    {
                        HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i, Encoding.GetEncoding("gb2312"));
                        HtmlNodeCollection provinceNodeList = htmlDoc.DocumentNode.SelectNodes("//tr[@class=\"provincetr\"]/td/a");
                        for (int j = 0; j < provinceNodeList.Count; j++)
                        {
                            HtmlNode provinceNode = provinceNodeList[j];
                            string hrefValue = provinceNode.GetAttributeValue("href", "");
                            int codeEndIndex = hrefValue.IndexOf(".");
                            string code = hrefValue.Substring(0, codeEndIndex);
                            string name = provinceNode.InnerText.Trim();

                            Dictionary<string, string> f2vs = new Dictionary<string, string>();
                            f2vs.Add("detailPageUrl", "http://www.stats.gov.cn/tjsj/tjbz/tjyqhdmhcxhfdm/" + year + "/" + hrefValue);
                            f2vs.Add("detailPageName", year + "_" + code);
                            f2vs.Add("year", year);
                            f2vs.Add("code", code);
                            f2vs.Add("name", name);
                            ew.AddRow(f2vs);
                        }
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
            }

            ew.SaveToDisk();
        }


        private ExcelWriter GetExcelWriter()
        {
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> columnDic = CommonUtil.InitStringIndexDic(new string[]{
                    "detailPageUrl",
                    "detailPageName", 
                    "cookie",
                    "grabStatus", 
                    "giveUpGrab",
                    "year",
                    "code",
                    "name"});

            string filePath = Path.Combine(exportDir, "统计用区划代码_Level2.xlsx");
            ExcelWriter ew = new ExcelWriter(filePath, "List", columnDic);
            return ew;
        }
    }
}
