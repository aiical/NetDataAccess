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
using NetDataAccess.Base.Writer;
using NetDataAccess.Base.Reader;
using NetDataAccess.Base.CsvHelper;
using NetDataAccess.Base.DB;
using HtmlAgilityPack;

namespace NetDataAccess.Extended.Yiguo
{
    public class LiangxianCategoryList : CustomProgramBase
    {
        public bool Run(string parameters, IListSheet listSheet)
        {
            return this.GenerateCategoryList(listSheet);
        }

        private bool GenerateCategoryList(IListSheet listSheet)
        {
            bool succeed = true;
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            Dictionary<string, int> categoryColumnDic = new Dictionary<string, int>();
            categoryColumnDic.Add("detailPageUrl", 0);
            categoryColumnDic.Add("detailPageName", 1);
            categoryColumnDic.Add("cookie", 2);
            categoryColumnDic.Add("grabStatus", 3);
            categoryColumnDic.Add("giveUpGrab", 4);
            categoryColumnDic.Add("category1Code", 5);
            categoryColumnDic.Add("category2Code", 6);
            categoryColumnDic.Add("category3Code", 7);
            categoryColumnDic.Add("category1Name", 8);
            categoryColumnDic.Add("category2Name", 9);
            categoryColumnDic.Add("category3Name", 10); 
            categoryColumnDic.Add("district", 11);
            string readDetailDir = this.RunPage.GetReadFileDir();
            string exportDir = this.RunPage.GetExportDir();
            string categoryFilePath = Path.Combine(exportDir, "两鲜网获取所有列表页.xlsx"); 
            ExcelWriter categoryCW = new ExcelWriter(categoryFilePath, "List", categoryColumnDic);

            GetList(listSheet, pageSourceDir, categoryCW);
            
            categoryCW.SaveToDisk(); 
            
            return succeed;
        }

        private void GetList(IListSheet listSheet, string pageSourceDir, ExcelWriter categoryCW)
        { 
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                string pageUrl = listSheet.PageUrlList[i];
                Dictionary<string, string> row = listSheet.GetRow(i);
                string cookie = row["cookie"];
                string district = row["detailPageName"];
                string localFilePath = this.RunPage.GetFilePath(pageUrl, pageSourceDir);
                TextReader tr = null;

                try
                {
                    tr = new StreamReader(localFilePath);
                    string webPageHtml = tr.ReadToEnd();
                    HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                    htmlDoc.LoadHtml(webPageHtml);
                    HtmlNodeCollection allCategory1Nodes = htmlDoc.DocumentNode.SelectNodes("//div[@class=\"all-sort-list\"]/div[@class=\"item bo\"]");
                    if (allCategory1Nodes != null)
                    { 
                        for (int j = 0; j < allCategory1Nodes.Count; j++)
                        {
                            HtmlNode category1Node = allCategory1Nodes[j];
                            HtmlNode c1Node = category1Node.SelectSingleNode("./h3/a");
                            string c1Name = CommonUtil.HtmlDecode(c1Node.InnerText.Trim());
                            string c1Url = c1Node.Attributes["href"].Value;
                            string[] c1UrlStrs = c1Url.Split(new string[] { "/", "." }, StringSplitOptions.RemoveEmptyEntries);
                            string c1Code = c1UrlStrs[c1UrlStrs.Length - 2];


                            //增加一级分类
                            {
                                Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                string listPageUrl = c1Url;
                                f2vs.Add("detailPageUrl", listPageUrl);
                                f2vs.Add("detailPageName", c1Name);
                                f2vs.Add("cookie", cookie);
                                f2vs.Add("category1Code", c1Code);
                                f2vs.Add("category2Code", "");
                                f2vs.Add("category3Code", "");
                                f2vs.Add("category1Name", c1Name);
                                f2vs.Add("category2Name", "");
                                f2vs.Add("category3Name", "");
                                f2vs.Add("district", district);
                                categoryCW.AddRow(f2vs);
                            }

                            HtmlNodeCollection allCategory2ParentNodes = category1Node.SelectNodes("./div[@class=\"item-list clearfix\"]/div[@class=\"subitem\"]");
                            if (allCategory2ParentNodes != null)
                            {
                                //第一个大类不要
                                for (int  k= 0; k < allCategory2ParentNodes.Count; k++)
                                { 
                                    HtmlNode category2ParentNode = allCategory2ParentNodes[k];
                                    HtmlNodeCollection allCategory2Nodes = category2ParentNode.SelectNodes("./h2");

                                    foreach (HtmlNode category2Node in allCategory2Nodes)
                                    {
                                        string childNodeText = category2Node.InnerText;
                                        if (childNodeText != "所有商品")
                                        {
                                            string c2Name = CommonUtil.HtmlDecode(category2Node.InnerText.Trim());
                                            HtmlNode c2Node = category2Node.SelectSingleNode("./a");
                                            string c2Code = "";
                                            string c2Url = "";
                                            if (c2Node != null)
                                            {
                                                c2Url = c2Node.Attributes["href"].Value;
                                                string[] c2UrlStrs = c2Url.Split(new string[] { "/", "." }, StringSplitOptions.RemoveEmptyEntries);
                                                c2Code = c2UrlStrs[c2UrlStrs.Length - 2];
                                            }

                                            //增加二级分类
                                            if (c2Url.Length != 0)
                                            {
                                                Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                                string listPageUrl = c2Url;
                                                f2vs.Add("detailPageUrl", listPageUrl);
                                                f2vs.Add("detailPageName", c1Name + "_" + c2Name);
                                                f2vs.Add("cookie", cookie);
                                                f2vs.Add("category1Code", c1Code);
                                                f2vs.Add("category2Code", c2Code);
                                                f2vs.Add("category3Code", "");
                                                f2vs.Add("category1Name", c1Name);
                                                f2vs.Add("category2Name", c2Name);
                                                f2vs.Add("category3Name", "");
                                                f2vs.Add("district", district);
                                                categoryCW.AddRow(f2vs);
                                            }

                                            List<HtmlNode> allCategory3Nodes = new List<HtmlNode>();
                                            HtmlNode nextNode = category2Node.NextSibling;
                                            while (nextNode != null && nextNode.Name.ToUpper() != "H2")
                                            {
                                                if (nextNode.Name.ToUpper() == "P")
                                                {
                                                    allCategory3Nodes.Add(nextNode);
                                                }
                                                nextNode = nextNode.NextSibling;
                                            }
                                            if (allCategory3Nodes.Count != 0)
                                            {
                                                //最末级为三级分类
                                                foreach (HtmlNode category3Node in allCategory3Nodes)
                                                {
                                                    HtmlNode c3Node = category3Node.SelectSingleNode("./a");
                                                    string c3Name = CommonUtil.HtmlDecode(c3Node.InnerText.Trim());
                                                    string c3Url = c3Node.Attributes["href"].Value;
                                                    string[] c3UrlStrs = c3Url.Split(new string[] { "/", "." }, StringSplitOptions.RemoveEmptyEntries);
                                                    string c3Code = c3UrlStrs[c3UrlStrs.Length - 2];

                                                    Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                                    string listPageUrl = c3Url;
                                                    f2vs.Add("detailPageUrl", listPageUrl);
                                                    f2vs.Add("detailPageName", c1Name + "_" + c2Name + "_" + c3Name);
                                                    f2vs.Add("cookie", cookie);
                                                    f2vs.Add("category1Code", c1Code);
                                                    f2vs.Add("category2Code", c2Code);
                                                    f2vs.Add("category3Code", c3Code);
                                                    f2vs.Add("category1Name", c1Name);
                                                    f2vs.Add("category2Name", c2Name);
                                                    f2vs.Add("category3Name", c3Name);
                                                    f2vs.Add("district", district);
                                                    categoryCW.AddRow(f2vs);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }                    
                }
                catch (Exception ex)
                {
                    this.RunPage.InvokeAppendLogText("读取出错. url = " + pageUrl + ". " + ex.Message, LogLevelType.Error, true);
                    throw ex;
                }
            }
        }
    }
}