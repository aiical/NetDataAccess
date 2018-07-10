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
    public class Level2 : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            this.GetCities(listSheet);
            return true;
        }

        private void GetCities(IListSheet listSheet)
        {
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            ExcelWriter ew = this.GetExcelWriter();

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"];
                string detailName = row["detailPageName"];
                string year = row["year"];
                string parentCode = row["code"];
                string parentName = row["name"];

                //添加父节点到下一级文件
                Dictionary<string, string> parentF2vs = new Dictionary<string, string>();
                parentF2vs.Add("detailPageUrl", detailUrl);
                parentF2vs.Add("detailPageName", detailName);
                parentF2vs.Add("year", year);
                parentF2vs.Add("code", parentCode);
                parentF2vs.Add("name", parentName);
                parentF2vs.Add("giveUpGrab", "Y");
                ew.AddRow(parentF2vs);

                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    try
                    {

                        HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i, Encoding.GetEncoding("gb2312"));
                        HtmlNodeCollection cityTypeNodeList = htmlDoc.DocumentNode.SelectNodes("//tr[@class=\"citytr\"]");
                        for (int j = 0; j < cityTypeNodeList.Count; j++)
                        {
                            HtmlNode cityTypeNode = cityTypeNodeList[j];
                            HtmlNodeCollection cityTypeFieldNodeList = cityTypeNode.SelectNodes("./td");
                            HtmlNode cityTypeCodeNode = cityTypeFieldNodeList[0];
                            HtmlNode cityTypeNameNode = cityTypeFieldNodeList[1];
                            string cityTypeCode = cityTypeCodeNode.InnerText.Trim();
                            string cityTypeName = cityTypeNameNode.InnerText.Trim();
                            HtmlNode linkNode = cityTypeCodeNode.SelectSingleNode("./a");

                            string hrefValue = "";
                            if (linkNode != null)
                            {
                                hrefValue = linkNode.GetAttributeValue("href", ""); 
                            }

                            Dictionary<string, string> f2vs = new Dictionary<string, string>();
                            f2vs.Add("detailPageUrl", "http://www.stats.gov.cn/tjsj/tjbz/tjyqhdmhcxhfdm/" + year + "/" + hrefValue);
                            f2vs.Add("detailPageName", year + "_" + cityTypeCode);
                            f2vs.Add("year", year);
                            f2vs.Add("code", cityTypeCode);
                            f2vs.Add("name", cityTypeName);
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

            string filePath = Path.Combine(exportDir, "统计用区划代码_Level3.xlsx");
            ExcelWriter ew = new ExcelWriter(filePath, "List", columnDic);
            return ew;
        }
    }
}
