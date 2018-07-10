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

namespace NetDataAccess.Extended.Jiaoyu.Shuyu
{
    public class GetSogouShuyuAllListPage : ExternalRunWebPage
    { 
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            try
            {
                this.GetList(listSheet);
                this.GetDetailList(listSheet); 
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return true;
        }

        private void GetList(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("cate1", 5);
            resultColumnDic.Add("cateId1", 6);
            resultColumnDic.Add("cate2", 7);
            resultColumnDic.Add("cateId2", 8);
            resultColumnDic.Add("cate3", 9);
            resultColumnDic.Add("cateId3", 10);
            resultColumnDic.Add("name", 11);
            string resultFilePath = Path.Combine(exportDir, "教育_术语_scel文件.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"];

                string cate1 = row["cate1"];
                string cate2 = row["cate2"];
                string cateId1 = row["cateId1"];
                string cateId2 = row["cateId2"];
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string localFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir);

                    try
                    {
                        HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                        HtmlNodeCollection itemNodes = htmlDoc.DocumentNode.SelectNodes("//div[@class=\"dict_detail_block\"]");

                        if (itemNodes != null)
                        {
                            for (int j = 0; j < itemNodes.Count; j++)
                            {
                                HtmlNode itemNode = itemNodes[j];
                                HtmlNode nameNode = itemNode.SelectSingleNode("./div[@class=\"dict_detail_title_block\"]/div[@class=\"detail_title\"]/a");
                                string itemName = nameNode.InnerText.Trim();

                                HtmlNode linkNode = itemNode.SelectSingleNode("./div[@class=\"dict_detail_show\"]/div[@class=\"dict_dl_btn\"]/a");
                                string linkUrl = linkNode.GetAttributeValue("href", "");

                                Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                f2vs.Add("detailPageUrl", linkUrl);
                                f2vs.Add("detailPageName", linkUrl);
                                f2vs.Add("cate1", row["cate1"]);
                                f2vs.Add("cateId1", row["cateId1"]);
                                f2vs.Add("cate2", row["cate2"]);
                                f2vs.Add("cateId2", row["cateId2"]);
                                f2vs.Add("cate3", row["cate3"]);
                                f2vs.Add("cateId3", row["cateId3"]);
                                f2vs.Add("name", itemName);
                                resultEW.AddRow(f2vs);
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
        private void GetDetailList(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("cate1", 5);
            resultColumnDic.Add("cateId1", 6);
            resultColumnDic.Add("cate2", 7);
            resultColumnDic.Add("cateId2", 8);
            resultColumnDic.Add("cate3", 9);
            resultColumnDic.Add("cateId3", 10);
            resultColumnDic.Add("name", 11);
            string resultFilePath = Path.Combine(exportDir, "教育_术语_scel详情.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"];

                string cate1 = row["cate1"];
                string cate2 = row["cate2"];
                string cateId1 = row["cateId1"];
                string cateId2 = row["cateId2"];
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string localFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir);

                    try
                    {
                        HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                        HtmlNodeCollection itemNodes = htmlDoc.DocumentNode.SelectNodes("//div[@class=\"dict_detail_block\"]");

                        if (itemNodes != null)
                        {
                            for (int j = 0; j < itemNodes.Count; j++)
                            {
                                HtmlNode itemNode = itemNodes[j];
                                HtmlNode nameNode = itemNode.SelectSingleNode("./div[@class=\"dict_detail_title_block\"]/div[@class=\"detail_title\"]/a");
                                string linkUrl = "https://pinyin.sogou.com" + nameNode.GetAttributeValue("href", "");
                                string itemName = nameNode.InnerText.Trim();

                                Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                f2vs.Add("detailPageUrl", linkUrl);
                                f2vs.Add("detailPageName", linkUrl);
                                f2vs.Add("cate1", row["cate1"]);
                                f2vs.Add("cateId1", row["cateId1"]);
                                f2vs.Add("cate2", row["cate2"]);
                                f2vs.Add("cateId2", row["cateId2"]);
                                f2vs.Add("cate3", row["cate3"]);
                                f2vs.Add("cateId3", row["cateId3"]);
                                f2vs.Add("name", itemName);
                                resultEW.AddRow(f2vs);
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

        private string getNameFromFullName(string fullCateName)
        {
            int splitBeginIndex = fullCateName.IndexOf("(");
            return splitBeginIndex < 0 ? fullCateName : fullCateName.Substring(0, splitBeginIndex);
        }

        private int getPageCountFromFullName(string fullName)
        {
            int splitBeginIndex = fullName.IndexOf("(");
            int splitEndIndex = fullName.IndexOf(")");
            return int.Parse(fullName.Substring(splitBeginIndex + 1, splitEndIndex - splitBeginIndex - 1)) / 10 + 1;
        }

    }
}