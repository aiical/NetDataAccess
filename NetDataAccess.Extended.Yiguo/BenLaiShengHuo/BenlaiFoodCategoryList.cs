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
    /// 本来生活
    /// 记录下商品分类及各个分类列表页首页URL
    /// </summary>
    public class BenlaiFoodCategoryList : CustomProgramBase
    {
        #region 入口函数
        /// <summary>
        /// 入口函数
        /// </summary>
        /// <param name="parameters">“扩展程序配置”信息中的parameters属性值</param>
        /// <param name="listSheet">输入文件，记录了要下载的所有页面</param>
        /// <returns></returns>
        public bool Run(string parameters, IListSheet listSheet)
        {
            return this.GetCategoryListFirstPageUrl(listSheet);
        }
        #endregion

        #region 从已经获取到的html中获取商品分类及各个分类列表页首页URL
        private bool GetCategoryListFirstPageUrl(IListSheet listSheet)
        { 
            //下载下来的网站首页保存到的目录（文件夹）
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            //输入文件所包含的列
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

            //输出的目录
            string exportDir = this.RunPage.GetExportDir();

            //输出文件夹名
            string categoryFilePath = Path.Combine(exportDir, "本来生活列表页首页.xlsx"); 

            //输出对象
            ExcelWriter categoryCW = new ExcelWriter(categoryFilePath, "List", categoryColumnDic);

            //解析html获取各分类列表页首页url
            GetCategoryListFirstPageUrl(listSheet, pageSourceDir, categoryCW);
            
            //保存到硬盘
            categoryCW.SaveToDisk(); 

            return true;
        }
        #endregion

        #region 解析html获取各分类列表页首页url
        private void GetCategoryListFirstPageUrl(IListSheet listSheet, string pageSourceDir, ExcelWriter categoryCW)
        {
            string urlPrefix ="http://www.benlai.com";
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
                    HtmlNodeCollection allCategory1Nodes = htmlDoc.DocumentNode.SelectNodes("//div[@class=\"tit_sort\"]/dl");
                    if (allCategory1Nodes != null)
                    {
                        for (int j = 0; j < allCategory1Nodes.Count; j++)
                        {
                            HtmlNode category1Node = allCategory1Nodes[j];
                            HtmlNode category1NameNode = category1Node.SelectSingleNode("./dt");
                            string category1Name = category1NameNode.InnerText.Trim();
                            string category1Code = "";
                            HtmlNodeCollection allCategory2Nodes = category1Node.SelectNodes("./dd[1]/div[1]/ul[1]/li");
                            for (int k = 0; k < allCategory2Nodes.Count; k++)
                            {
                                HtmlNode category2Node = allCategory2Nodes[k];
                                HtmlNode category2NameNode = category2Node.SelectSingleNode("./div[1]");
                                string category2Name = category2NameNode.InnerText.Trim();
                                string category2Code = "";
                                HtmlNodeCollection allCategory3Nodes = category2Node.SelectNodes("./div[2]/em");
                                for (int l = 0; l <= allCategory3Nodes.Count; l++)
                                {
                                    string category3Name = "";
                                    string category3Code = "";
                                    string url = "";
                                    if (l < allCategory3Nodes.Count)
                                    {
                                        HtmlNode category3Node = allCategory3Nodes[l];
                                        HtmlNode category3NameNode = category3Node.SelectSingleNode("./a");
                                        category3Name = category3NameNode.InnerText.Trim(); 

                                        //url格式为/list-36-1176-1183.html
                                        url = category3NameNode.Attributes["href"].Value;
                                        string[] urlSplits = url.Split(new string[] { "-", "." }, StringSplitOptions.RemoveEmptyEntries);
                                        category1Code = urlSplits[1];
                                        category2Code = urlSplits[2];
                                        category3Code = urlSplits[3];
                                    }
                                    else
                                    {
                                        url = category2NameNode.SelectSingleNode("./a").Attributes["href"].Value;
                                    }
                                    string categoryName = category1Name + "_" + category2Name + "_" + category3Name;
                                     
                                    Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                    f2vs.Add("detailPageUrl", urlPrefix + url);
                                    f2vs.Add("detailPageName", district + "_" + categoryName + "_" + category3Code);
                                    f2vs.Add("cookie", cookie);
                                    f2vs.Add("category1Code", category1Code);
                                    f2vs.Add("category2Code", category2Code);
                                    f2vs.Add("category3Code", category3Code);
                                    f2vs.Add("category1Name", category1Name);
                                    f2vs.Add("category2Name", category2Name);
                                    f2vs.Add("category3Name", category3Name);
                                    f2vs.Add("district", district);
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