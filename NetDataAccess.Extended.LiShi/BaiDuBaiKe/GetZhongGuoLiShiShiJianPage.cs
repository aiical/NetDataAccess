using HtmlAgilityPack;
using NetDataAccess.Base.Common;
using NetDataAccess.Base.DB;
using NetDataAccess.Base.UI;
using NetDataAccess.Base.Writer;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace NetDataAccess.Extended.LiShi.BaiDuBaiKe
{
    public class GetZhongGuoLiShiShiJianPage : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            this.GetShiJianPageUrls(listSheet);
            return true;
        }

        private void GetShiJianPageUrls(IListSheet listSheet)
        {
            ExcelWriter resultEW= this.CreateResultListWriter();

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                HtmlNodeCollection level2Nodes = htmlDoc.DocumentNode.SelectNodes("//div[@class=\"para-title level-2\"]");
                foreach (HtmlNode level2Node in level2Nodes)
                {
                    HtmlNode shiDaiL2Node = level2Node.SelectSingleNode("./h2");
                    string shiDaiL2 = "";
                    foreach (HtmlNode l2TitleNode in shiDaiL2Node.ChildNodes)
                    {
                        if (l2TitleNode.Name.ToLower() != "span")
                        {
                            string text = CommonUtil.HtmlDecode(l2TitleNode.InnerText).Trim();
                            if (text.Length > 0)
                            {
                                shiDaiL2 = text;
                                break;
                            }
                        }
                    }

                    HtmlNode nextNode = level2Node.NextSibling;
                    string shiDaiL3 = "";

                    while (nextNode != null)
                    {
                        string className = nextNode.GetAttributeValue("class", "");
                        if (className == "para-title level-2")
                        {
                            break;
                        }
                        else if (className == "para-title level-3")
                        {
                            HtmlNode shiDaiL3Node = nextNode.SelectSingleNode("./h3");
                            foreach (HtmlNode l3TitleNode in shiDaiL3Node.ChildNodes)
                            {
                                if (l3TitleNode.Name.ToLower() != "span")
                                {
                                    string text = CommonUtil.HtmlDecode(l3TitleNode.InnerText).Trim();
                                    if (text.Length > 0)
                                    {
                                        shiDaiL3 = text;
                                        break;
                                    }
                                }
                            }
                        }
                        else if (className == "para")
                        {
                            HtmlNodeCollection childNodes = nextNode.ChildNodes;
                            for (int j = 0; j < childNodes.Count; j++)
                            {
                                HtmlNode shiJianNode = childNodes[j];
                                string shiJianName = CommonUtil.HtmlDecode(shiJianNode.InnerText).Trim();
                                if (shiJianName.Length != 0 && shiJianName != "·")
                                {
                                    string url = shiJianNode.Name.ToLower() == "a" ? ("https://baike.baidu.com" + shiJianNode.GetAttributeValue("href", "")) : "";

                                    Dictionary<string, string> resultRow = new Dictionary<string, string>();
                                    resultRow.Add("detailPageUrl", url);
                                    resultRow.Add("detailPageName", url);
                                    resultRow.Add("shiJian", shiJianName);
                                    resultRow.Add("shiDaiL2", shiDaiL2);
                                    resultRow.Add("shiDaiL3", shiDaiL3);
                                    resultEW.AddRow(resultRow);
                                }
                            }
                        }
                        nextNode = nextNode.NextSibling;
                    }
                }
            }
            resultEW.SaveToDisk();
        }
        private ExcelWriter CreateResultListWriter()
        {
            String exportDir = this.RunPage.GetExportDir();
            string resultFilePath = Path.Combine(exportDir, "百度百科_中国历史事件_事件页面.xlsx");

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("shiJian", 5); 
            resultColumnDic.Add("shiDaiL2", 6);
            resultColumnDic.Add("shiDaiL3", 7);
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }
    }
}
