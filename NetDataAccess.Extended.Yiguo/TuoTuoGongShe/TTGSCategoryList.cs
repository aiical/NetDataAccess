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
    public class TTGSCategoryList : CustomProgramBase
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
            string categoryFilePath = Path.Combine(exportDir, "沱沱工社列表页首页.xlsx"); 
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
                    HtmlNodeCollection allCategory1Nodes = htmlDoc.DocumentNode.SelectNodes("//div[@class=\"tall_menu\"]/div[@class=\"tall_menu_list\"]/ul/li");
                    List<string> c1Names = new List<string>();
                    List<string> c1Codes = new List<string>();
                    if (allCategory1Nodes != null)
                    {
                        for (int j = 0; j < allCategory1Nodes.Count; j++)
                        {
                            HtmlNode category1Node = allCategory1Nodes[j];
                            HtmlNode c1Node = category1Node.SelectSingleNode("./div[1]/span[1]/a");
                            string c1Code = c1Node.Attributes["cat"].Value;
                            string c1Name = c1Node.InnerText.Trim();
                            c1Codes.Add(c1Code);
                            c1Names.Add(c1Name);
                        }
                    }

                    HtmlNodeCollection allCategory2ParentNodes = htmlDoc.DocumentNode.SelectNodes("//div[@class=\"tall_menu_counts\"]/div[@class=\"tall_menu_count\"]/div[@class=\"tall_menu_clist\"]/ul");
                    if (allCategory2ParentNodes != null)
                    {
                        for (int j = 0; j < allCategory2ParentNodes.Count; j++)
                        {
                            string c1Code = c1Codes[j];
                            string c1Name = c1Names[j];

                            HtmlNode category2ParentNode = allCategory2ParentNodes[j];
                            HtmlNodeCollection allCategory2Nodes = category2ParentNode.SelectNodes("./li");
                            if (allCategory2Nodes != null)
                            {
                                for (int k = 0; k < allCategory2Nodes.Count; k++)
                                {
                                    HtmlNode category2Node = allCategory2Nodes[k];
                                    HtmlNode c2Node = category2Node.SelectSingleNode("./span[@class=\"span_left\"]/a");
                                    string c2Code = c2Node.Attributes["cat"].Value;
                                    string c2Name = c2Node.InnerText.Trim();

                                    HtmlNodeCollection allCategory3Nodes = category2Node.SelectNodes("./span[@class=\"span_right\"]/a");
                                    for (int m = 0; m < allCategory3Nodes.Count; m++)
                                    {
                                        HtmlNode c3Node = allCategory3Nodes[m];
                                        string c3Name = c3Node.InnerText.Trim();
                                        string c3Code = c3Node.Attributes["cat"].Value;
                                        string c3Url = c3Node.Attributes["href"].Value;

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
                catch (Exception ex)
                {
                    this.RunPage.InvokeAppendLogText("读取出错. url = " + pageUrl + ". " + ex.Message, LogLevelType.Error, true);
                    throw ex;
                }
            }
        }
    }
}