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
    public class FeiniuTypeList : CustomProgramBase
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
            string categoryFilePath = Path.Combine(exportDir, this.RunPage.Project.Name + "_EveryCategoryFirstPage.xlsx"); 
            ExcelWriter categoryCW = new ExcelWriter(categoryFilePath, "List", categoryColumnDic);

            GetList(listSheet, pageSourceDir, categoryCW);
            
            categoryCW.SaveToDisk(); 

            return succeed;
        }

        private void GetList(IListSheet listSheet, string pageSourceDir, ExcelWriter categoryCW)
        {
            string urlPrefix = "http://www.feiniu.com/";
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                string pageUrl = listSheet.PageUrlList[i];
                Dictionary<string, string> row = listSheet.GetRow(i); 
                string district = row["detailPageName"];
                string localFilePath = this.RunPage.GetFilePath(pageUrl, pageSourceDir);
                TextReader tr = null;

                try
                {
                    tr = new StreamReader(localFilePath);
                    string webPageHtml = tr.ReadToEnd();
                    HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                    htmlDoc.LoadHtml(webPageHtml);
                    HtmlNodeCollection allCategory1Nodes = htmlDoc.DocumentNode.SelectNodes("//dl[@class=\"sitemap-catagory\"]");
                    if (allCategory1Nodes != null)
                    {
                        HtmlNode category1Node = allCategory1Nodes[0];
                        HtmlNode category1NameNode = category1Node.SelectSingleNode("./dt[1]/a[1]");
                        string category1Name = category1NameNode.InnerText.Trim();
                        string category1Code = category1NameNode.Attributes["id"].Value;
                        HtmlNodeCollection allCategory2Nodes = category1Node.SelectNodes("./dd");
                        if (allCategory2Nodes != null)
                        {
                            for (int k = 0; k < allCategory2Nodes.Count; k++)
                            {
                                HtmlNode category2Node = allCategory2Nodes[k];
                                HtmlNode category2NameNode = category2Node.SelectSingleNode("./dl[1]/dt[1]/a[1]");
                                string category2Name = category2NameNode.InnerText.Trim();
                                string category2Url = category2NameNode.Attributes["href"].Value;
                                string category2Code = category2Url.Substring(category2Url.LastIndexOf("/") + 1);
                                HtmlNodeCollection allCategory3Nodes = category2Node.SelectNodes("./dl[1]/dd[1]/a");
                                if (allCategory3Nodes != null)
                                {
                                    for (int l = 0; l < allCategory3Nodes.Count; l++)
                                    {
                                        HtmlNode category3Node = allCategory3Nodes[l];
                                        string category3Name = category3Node.InnerText.Trim();
                                        string category3Url = category3Node.Attributes["href"].Value;
                                        string category3Code = category3Url.Substring(category3Url.LastIndexOf("/") + 1);

                                        string categoryName = category1Name + "_" + category2Name + "_" + category3Name;

                                        //huadong
                                        Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                        f2vs.Add("detailPageUrl", urlPrefix + category3Url);
                                        f2vs.Add("detailPageName", district + "_" + categoryName + "_" + category3Code); 
                                        f2vs.Add("category1Code", category1Code);
                                        f2vs.Add("category2Code", category2Code);
                                        f2vs.Add("category3Code", category3Code);
                                        f2vs.Add("category1Name", category1Name);
                                        f2vs.Add("category2Name", category2Name);
                                        f2vs.Add("category3Name", category3Name);
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