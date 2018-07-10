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

namespace NetDataAccess.Extended.Tanyin
{
    /// <summary>
    /// Tangongye展会
    /// 运行此程序前，系统已经提前爬取了listSheet中指定的展会详情页html，
    /// 然后此扩展程序解析这些详情页，形成结构化信息存储在excel中，
    /// 并从详情信息html获取展会信息正文中的img图片url
    /// </summary>
    public class TangongyeZhanhuiGetDetail : CustomProgramBase
    {
        #region 入口函数
        /// <summary>
        /// 入口函数
        /// </summary>
        /// <param name="parameters">“扩展程序配置”信息中的parameters属性值</param>
        /// <param name="listSheet">输入文件，记录了要下载的所有页面（其中包含的记录，再爬取前先由excel导入到sqlite表，然后系统操作的一直是sqlite表里的记录）</param>
        /// <returns></returns>
        public bool Run(string parameters, IListSheet listSheet)
        {
            return GetAllPageDetailInfo(listSheet) && GetAllPageImgInfo(listSheet);
        }
        #endregion

        #region 逐个详情页获取展会信息
        /// <summary>
        /// 逐个详情页获取展会信息
        /// </summary>
        /// <param name="listSheet"></param>
        /// <returns></returns>
        private bool GetAllPageDetailInfo(IListSheet listSheet)
        {
            //输出目录（从配置中获取）
            string exportDir = this.RunPage.GetExportDir();

            //下载下来的详情页html保存目录（文件夹）
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            //输出excel表格包含的列，此文件提供给客户
            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                "标题",
                "编码", 
                "日期",
                "发布方", 
                "url",
                "正文HTML"});

            //输出文件地址
            string resultFilePath = Path.Combine(exportDir, "展会-Tangongye展会.xlsx");

            //输出文件对象
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;
            string detailPageNameColumnName = SysConfig.DetailPageNameFieldName;

            //循环输入文件中的所有行
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);

                //如果此行没有放弃爬取（爬取工具可以配置成爬取失败后放弃爬取，被放弃爬取的行记录到sqlite表中）
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string url = row[detailPageUrlColumnName];
                    string code = row[detailPageNameColumnName];
                    string title = row["title"].Trim();
                    string date = row["date"].Trim();

                    //下载下来的html地址
                    string localFilePath = this.RunPage.GetFilePath(url, pageSourceDir);
                    TextReader tr = null;

                    try
                    {
                        //读取展会详情页html，并加载到HtmlDocument对象中（系统提供了构造HtmlDocument对象的方法，写这段代码的时候忘记用了）
                        tr = new StreamReader(localFilePath);
                        string webPageHtml = tr.ReadToEnd();
                        HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                        htmlDoc.LoadHtml(webPageHtml);

                        //标题在解析列表页的时候已经获取了
                        //HtmlNode titleNode = htmlDoc.DocumentNode.SelectSingleNode("//span[@id=\"labTitle\"]"); 

                        //从HtmlDocument对象中获取作者
                        HtmlNode authorNode = htmlDoc.DocumentNode.SelectSingleNode("//span[@id=\"labAuthor\"]");
                        string author = authorNode == null ? "" : CommonUtil.HtmlDecode(authorNode.InnerText).Trim();

                        //HtmlDocument对象中获取展会信息正文HTML，注意这里是取的InnerHtml
                        HtmlNode contentNode = htmlDoc.DocumentNode.SelectSingleNode("//div[@id=\"newconcent\"]");
                        string contextHtml = contentNode.InnerHtml;

                        //保存信息到resultEW（ExcelWriter对象）
                        Dictionary<string, object> f2vs = new Dictionary<string, object>();
                        f2vs.Add("标题", title);
                        f2vs.Add("编码", code);
                        f2vs.Add("日期", date);
                        f2vs.Add("发布方", author);
                        f2vs.Add("url", url);
                        f2vs.Add("正文HTML", contextHtml);
                        resultEW.AddRow(f2vs);

                    }
                    catch (Exception ex)
                    {
                        if (tr != null)
                        {
                            tr.Dispose();
                            tr = null;
                        }
                        this.RunPage.InvokeAppendLogText("读取出错.  " + ex.Message + " LocalPath = " + localFilePath, LogLevelType.Error, true);
                        throw ex;
                    }
                }
            }

            //输出到硬盘
            resultEW.SaveToDisk();

            return true;
        }
        #endregion

        #region 获取展会详情信息中用到图片的地址
        /// <summary>
        /// 获取展会详情信息中用到图片的地址
        /// </summary>
        /// <param name="listSheet"></param>
        /// <returns></returns>
        private bool GetAllPageImgInfo(IListSheet listSheet)
        {
            string exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                "detailPageUrl",
                "detailPageName", 
                "cookie",
                "grabStatus", 
                "giveUpGrab",
                "pageCode"});

            //此输出文件是项目“Tangongye展会获取图片”的输入文件
            string resultFilePath = Path.Combine(exportDir, "Tangongye展会获取图片.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;
            string detailPageNameColumnName = SysConfig.DetailPageNameFieldName;
            Dictionary<string, string> allImageNames = new Dictionary<string, string>();

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string url = row[detailPageUrlColumnName];
                    string code = row[detailPageNameColumnName]; 
                    string localFilePath = this.RunPage.GetFilePath(url, pageSourceDir);
                    TextReader tr = null;

                    try
                    {
                        tr = new StreamReader(localFilePath);
                        string webPageHtml = tr.ReadToEnd();

                        HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                        htmlDoc.LoadHtml(webPageHtml); 
                        
                        HtmlNode contentNode = htmlDoc.DocumentNode.SelectSingleNode("//div[@id=\"newconcent\"]");
                        List<HtmlNode> allImgNodes = new List<HtmlNode>();

                        this.GetAllImgNodes(contentNode, allImgNodes);

                        foreach (HtmlNode imgNode in allImgNodes)
                        {
                            string imgUrl = imgNode.Attributes["src"].Value;
                            Dictionary<string, object> f2vs = new Dictionary<string, object>();
                            string name = code + "_" + imgUrl;
                            if (!allImageNames.ContainsKey(name))
                            {
                                allImageNames.Add(name, null);
                                f2vs.Add("pageCode", code);
                                f2vs.Add("detailPageUrl", imgUrl);
                                f2vs.Add("detailPageName", name);
                                resultEW.AddRow(f2vs);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        if (tr != null)
                        {
                            tr.Dispose();
                            tr = null;
                        }
                        this.RunPage.InvokeAppendLogText("读取出错.  " + ex.Message + " LocalPath = " + localFilePath, LogLevelType.Error, true);
                        throw ex;
                    }
                }
            }
            resultEW.SaveToDisk();

            return true;
        }
        #endregion

        #region 递归获取子元素包含的所有IMG元素
        private void GetAllImgNodes(HtmlNode parentNode, List<HtmlNode> allImgNodes)
        {
            HtmlNodeCollection allChildNodes = parentNode.ChildNodes;
            if (allChildNodes != null)
            {
                foreach (HtmlNode childNode in allChildNodes)
                {
                    if (childNode.Name.ToLower() == "img")
                    {
                        allImgNodes.Add(childNode);
                    }
                    else
                    {
                        GetAllImgNodes(childNode, allImgNodes);
                    }
                }
            }
        }
        #endregion
    }
}