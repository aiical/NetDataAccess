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
    public class WomaiCategoryList : CustomProgramBase
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
            categoryColumnDic.Add("category1Name", 7);
            categoryColumnDic.Add("category2Name", 8); 
            categoryColumnDic.Add("district", 9);
            string readDetailDir = this.RunPage.GetReadFileDir();
            string exportDir = this.RunPage.GetExportDir();
            string categoryFilePath = Path.Combine(exportDir, "我买网列表页首页.xlsx"); 
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
                    tr = new StreamReader(localFilePath,Encoding.GetEncoding("GBK"));
                    string webPageHtml = tr.ReadToEnd();
                    HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                    htmlDoc.LoadHtml(webPageHtml);
                    HtmlNodeCollection allCategoryMainNodes = htmlDoc.DocumentNode.SelectNodes("//li[@class=\"kinds\"]/div[@class=\"sub_kinds\"]/div[@class=\"kinds-box viewport\"]/div[@class=\"c_kinds overview\"]");
                    if (allCategoryMainNodes != null)
                    {
                        for (int j = 0; j < allCategoryMainNodes.Count; j++)
                        {
                            HtmlNode categoryMainNode = allCategoryMainNodes[j];
                            HtmlNodeCollection allCategory1Nodes = categoryMainNode.SelectNodes("./h4[@class=\"sub_head\"]/a");
                            List<string> c1CodeList = new List<string>();
                            List<string> c1NameList = new List<string>();
                            if (allCategory1Nodes != null)
                            {
                                for (int k = 0; k < allCategory1Nodes.Count; k++)
                                {
                                    HtmlNode category1Node = allCategory1Nodes[k];
                                    string category1Name = category1Node.InnerText.Trim();
                                    string category1Url = category1Node.Attributes["href"].Value;
                                    string[] c1UrlSplits = category1Url.Split(new string[] { "-", "." }, StringSplitOptions.RemoveEmptyEntries);
                                    string category1Code = c1UrlSplits[c1UrlSplits.Length - 2];
                                    c1CodeList.Add(category1Code);
                                    c1NameList.Add(category1Name);
                                }
                            }

                            HtmlNodeCollection allCategory2ParentNodes = categoryMainNode.SelectNodes("./ul[@class=\"sub_cont\"]");
                            if (allCategory2ParentNodes != null)
                            { 
                                for (int l = 0; l < allCategory2ParentNodes.Count; l++)
                                {
                                    string category1Code = c1CodeList[l];
                                    string category1Name = c1NameList[l];

                                    HtmlNode categor2ParentNode = allCategory2ParentNodes[l];
                                    HtmlNodeCollection allCategory2Nodes = categor2ParentNode.SelectNodes("./li[@class=\"sub_kind\"]/a");
                                    for (int m = 0; m < allCategory2Nodes.Count; m++)
                                    {
                                        HtmlNode category2Node = allCategory2Nodes[m];
                                        string category2Name = category2Node.InnerText.Trim();
                                        string category2Url = category2Node.Attributes["href"].Value;
                                        string[] c2UrlSplits = category2Url.Split(new string[] { "-", "." }, StringSplitOptions.RemoveEmptyEntries);
                                        string category2Code = c2UrlSplits[c2UrlSplits.Length - 2];


                                        string categoryName = category1Name + "_" + category2Name;

                                        Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                        string listPageUrl = "http://www.womai.com/ProductList.do?mid=0&Cid=" + category2Code + "&mainColumnId=&page=1&brand=-1&rypId=608&zhId=605&orderBy=&isPromotions=&sellable=&Keywords=&Keyword=&isKeyCommendClick=1&sellerid=&selAttr=&selCol=&urllist=";
                                        f2vs.Add("detailPageUrl", listPageUrl);
                                        f2vs.Add("detailPageName", categoryName + "_" + category2Code);
                                        f2vs.Add("cookie", cookie);
                                        f2vs.Add("category1Code", category1Code);
                                        f2vs.Add("category2Code", category2Code);
                                        f2vs.Add("category1Name", category1Name);
                                        f2vs.Add("category2Name", category2Name);
                                        f2vs.Add("district", district);
                                        categoryCW.AddRow(f2vs);
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