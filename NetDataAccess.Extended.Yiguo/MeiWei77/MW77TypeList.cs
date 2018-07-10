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
using NetDataAccess.Base.Server;

namespace NetDataAccess.Extended.Yiguo
{
    /// <summary>
    /// 美味七七
    /// 记录下商品分类及各个分类列表页首页URL
    /// </summary>
    public class MW77TypeList : CustomProgramBase
    {
        #region 入口函数
        public bool Run(string parameters, IListSheet listSheet)
        {
            return this.GenerateFoodCategoryList(listSheet);
        }
        #endregion

        #region 从已经获取到的html中获取商品分类及各个分类列表页首页URL
        private bool GenerateFoodCategoryList(IListSheet listSheet)
        {
            bool succeed = true;
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            string[] resultColumns = new string[]{ "detailPageUrl",
                "detailPageName",
                "cookie",
                "grabStatus",
                "giveUpGrab", 
                "category1Code", 
                "category2Code",
                "category3Code", 
                "category1Name", 
                "category2Name",
                "category3Name"};
            Dictionary<string, int> categoryColumnDic = CommonUtil.InitStringIndexDic(resultColumns);
            string readDetailDir = this.RunPage.GetReadFileDir();
            string exportDir = this.RunPage.GetExportDir();
            string categoryFilePath = Path.Combine(exportDir, this.RunPage.Project.Name + "_EveryCategoryFirstPage.xlsx"); 
            ExcelWriter categoryCW = new ExcelWriter(categoryFilePath, "List", categoryColumnDic);

            GenerateFoodCategoryList(listSheet, pageSourceDir, categoryCW, "http://www.yummy77.com"); 

            categoryCW.SaveToDisk();

            //执行后续任务
            TaskManager.StartTask("易果", "美味77列表页首页", categoryFilePath, null, null, false);
            
            return succeed;
        }
        #endregion

        #region 从已经获取到的html中获取商品分类及各个分类列表页首页URL
        private void GenerateFoodCategoryList(IListSheet listSheet, string pageSourceDir, ExcelWriter categoryCW, string urlPrefix)
        {
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                string pageUrl = listSheet.PageUrlList[i];
                string localFilePath = this.RunPage.GetFilePath(pageUrl, pageSourceDir);
                TextReader tr = null;

                try
                {
                    tr = new StreamReader(localFilePath);
                    string webPageHtml = tr.ReadToEnd();
                    HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                    htmlDoc.LoadHtml(webPageHtml);
                    HtmlNodeCollection allCategory1Nodes = htmlDoc.DocumentNode.SelectNodes("//*[@id=\"main_cat\"]/li");
                    if (allCategory1Nodes != null)
                    {
                        for (int j = 0; j < allCategory1Nodes.Count; j++)
                        {
                            HtmlNode category1Node = allCategory1Nodes[j];
                            HtmlNode category1NameNode = category1Node.SelectSingleNode("./div[1]/a");
                            string category1Name = category1NameNode.InnerText.Trim();
                            string category1Href = category1NameNode.Attributes["href"].Value;
                            string category1Code = category1Href.Split(new string[] { "/", "." }, StringSplitOptions.RemoveEmptyEntries)[2];
                            HtmlNodeCollection allCategory2Nodes = category1Node.SelectNodes("./div[2]/div[1]/dl");
                            for (int k = 0; k < allCategory2Nodes.Count; k++)
                            {
                                HtmlNode category2Node = allCategory2Nodes[k];
                                HtmlNode category2NameNode = category2Node.SelectSingleNode("./dt[1]/a");
                                string category2Name = category2NameNode.InnerText.Trim();
                                string category2Href = category2NameNode.Attributes["href"].Value;
                                string category2Code = category2Href.Split(new string[] { "/", "." }, StringSplitOptions.RemoveEmptyEntries)[1];
                                HtmlNodeCollection allCategory3Nodes = category2Node.SelectNodes("./dd/a");
                                for (int l = 0; l < allCategory3Nodes.Count; l++)
                                {
                                    HtmlNode category3NameNode = allCategory3Nodes[l]; 
                                    string category3Name = category3NameNode.InnerText.Trim();
                                    string category3Href = category3NameNode.Attributes["href"].Value;
                                    string category3Code = category3Href.Split(new string[] { "/", "." }, StringSplitOptions.RemoveEmptyEntries)[1];

                                    string categoryName = category1Name + "_" + category2Name + "_" + category3Name;

                                    //huadong
                                    Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                    f2vs.Add("detailPageUrl", urlPrefix + category3Href);
                                    f2vs.Add("detailPageName", categoryName);
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
                catch (Exception ex)
                {
                    this.RunPage.InvokeAppendLogText("读取出错. url = " + pageUrl + ". " + ex.Message, LogLevelType.Error, true);
                    throw ex;
                }
            }
        }
        #endregion
    }
}