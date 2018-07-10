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
    public class GetSogouShuyuXiaoleiListPage : ExternalRunWebPage
    { 
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            try
            {
                this.GetList(listSheet); 
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
            resultColumnDic.Add("pageIndex", 11);
            string resultFilePath = Path.Combine(exportDir, "教育_术语_列表.xlsx");
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
                        HtmlNodeCollection cate2NoChildLinkNodes = htmlDoc.DocumentNode.SelectNodes("//table[@class=\"cate_words_list\"]/tbody/tr/td/div[contains(@class, \"cate_no_child\")]/a");

                        if (cate2NoChildLinkNodes != null)
                        {
                            for (int j = 0; j < cate2NoChildLinkNodes.Count; j++)
                            {
                                HtmlNode cate2NoChildLinkNode = cate2NoChildLinkNodes[j];
                                string linkUrl = cate2NoChildLinkNode.GetAttributeValue("href", "");
                                string cateName = cate2NoChildLinkNode.InnerText.Trim();
                                int linkIdBeginIndex = linkUrl.LastIndexOf("/") + 1;
                                string id = linkUrl.Substring(linkIdBeginIndex).Trim();

                                int pageCount = this.getPageCountFromFullName(cateName);
                                for (int k = 0; k < pageCount; k++)
                                {
                                    if (cate2 == null || cate2.Length == 0)
                                    {
                                        string newUrl = "https://pinyin.sogou.com" + linkUrl + "/default/" + (k + 1).ToString();
                                        Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                        f2vs.Add("detailPageUrl", newUrl);
                                        f2vs.Add("detailPageName", newUrl);
                                        f2vs.Add("cate1", cate1);
                                        f2vs.Add("cateId1", cateId1);
                                        f2vs.Add("cate2", this.getNameFromFullName(cateName));
                                        f2vs.Add("cateId2", id);
                                        f2vs.Add("cate3", this.getNameFromFullName(cateName));
                                        f2vs.Add("cateId3", id);
                                        f2vs.Add("pageIndex", (k + 1).ToString());
                                        resultEW.AddRow(f2vs);
                                    }
                                    else
                                    {
                                        string newUrl = "https://pinyin.sogou.com" + linkUrl + "/default/" + (k + 1).ToString();
                                        Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                        f2vs.Add("detailPageUrl", newUrl);
                                        f2vs.Add("detailPageName", newUrl);
                                        f2vs.Add("cate1", cate1);
                                        f2vs.Add("cateId1", cateId1);
                                        f2vs.Add("cate2", this.getNameFromFullName(cate2));
                                        f2vs.Add("cateId2", cateId2);
                                        f2vs.Add("cate3", this.getNameFromFullName(cateName));
                                        f2vs.Add("cateId3", id);
                                        f2vs.Add("pageIndex", (k + 1).ToString());
                                        resultEW.AddRow(f2vs);
                                    }
                                }
                            }
                        }


                        HtmlNodeCollection cate2HasChildDivNodes = htmlDoc.DocumentNode.SelectNodes("//table[@class=\"cate_words_list\"]/tbody/tr/td/div[contains(@class, \"cate_has_child\")]");

                        if (cate2HasChildDivNodes != null)
                        {
                            for (int j = 0; j < cate2HasChildDivNodes.Count; j++)
                            {
                                HtmlNode cate2HasChildNode = cate2HasChildDivNodes[j];
                                HtmlNode cate2HasChildLinkNode = cate2HasChildNode.SelectSingleNode("./a");
                                string linkCate2Url = cate2HasChildLinkNode.GetAttributeValue("href", "");
                                cate2 = cate2HasChildLinkNode.InnerText.Trim();
                                int linkCate2IdBeginIndex = linkCate2Url.LastIndexOf("/") + 1;
                                cateId2 = linkCate2Url.Substring(linkCate2IdBeginIndex).Trim();

                                HtmlNode tempNode = cate2HasChildNode.NextSibling;
                                while (tempNode.GetAttributeValue("class", "") != "cate_children_show")
                                {
                                    tempNode = tempNode.NextSibling;
                                }

                                if (tempNode == null)
                                {
                                    throw new Exception("没有明细分类");
                                }
                                else
                                {
                                    HtmlNodeCollection childNodes = tempNode.SelectNodes("./table/tbody/tr/td/div[@class=\"cate_child_name\"]/a");

                                    if (childNodes == null || childNodes.Count == 0)
                                    {
                                        throw new Exception("没有明细分类链接");
                                    }
                                    else
                                    {

                                        foreach (HtmlNode childNode in childNodes)
                                        {
                                            string linkCate3Url = childNode.GetAttributeValue("href", "");
                                            string cate3 = childNode.InnerText.Trim();
                                            int linkCate3IdBeginIndex = linkCate3Url.LastIndexOf("/") + 1;
                                            string cateId3 = linkCate3Url.Substring(linkCate3IdBeginIndex).Trim();

                                            int pageCount = this.getPageCountFromFullName(cate3);
                                            for (int k = 0; k < pageCount; k++)
                                            {
                                                string newUrl = "https://pinyin.sogou.com" + linkCate3Url + "/default/" + (k + 1).ToString();
                                                Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                                f2vs.Add("detailPageUrl", newUrl);
                                                f2vs.Add("detailPageName", newUrl);
                                                f2vs.Add("cate1", cate1);
                                                f2vs.Add("cateId1", cateId1);
                                                f2vs.Add("cate2", this.getNameFromFullName(cate2));
                                                f2vs.Add("cateId2", cateId2);
                                                f2vs.Add("cate3", this.getNameFromFullName(cate3));
                                                f2vs.Add("cateId3", cateId3);
                                                f2vs.Add("pageIndex", (k + 1).ToString());
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