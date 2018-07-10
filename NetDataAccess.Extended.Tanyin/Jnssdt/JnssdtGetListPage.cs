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
    /// 节能搜索低碳获取列表页
    /// 运行此程序前，系统已经提前爬取了listSheet中指定的节能搜索低碳获取列表页html
    /// 利用此程序，获取到节能搜索中搜索低碳一共多少页符合条件的页面，并按照url规则形成所有页面的url，为“节能搜索低碳获取列表页”提供下载列表
    /// </summary>
    public class JnssdtGetListPage : CustomProgramBase
    {
        #region 入口函数
        /// <summary>
        /// 入口函数
        /// </summary>
        /// <param name="parameters">“扩展程序配置”信息中的parameters属性值</param>
        /// <param name="listSheet">输入文件，记录了要下载的所有页面（其中包含的记录，再爬取前先由excel导入到sqlite表，然后系统操作的一直是sqlite表里的记录，此项目的仅一条记录）</param>
        /// <returns></returns>
        public bool Run(string parameters, IListSheet listSheet)
        {
            return GetAllListPageUrl(parameters, listSheet);
        }
        #endregion 

        #region 获取所有列表页url
        private bool GetAllListPageUrl(string parameters, IListSheet listSheet)
        {
            //配置文件中，扩展程序指定的parameters属性值
            string[] parameterArray = parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);

            //需要抓取此日期后的展会信息
            string toDateStr = parameterArray[0];
            DateTime toDate = DateTime.ParseExact(toDateStr, "yyyy-MM-dd", System.Globalization.CultureInfo.CurrentCulture);

            //输出目录（从配置中获取）
            string exportDir = this.RunPage.GetExportDir();

            //已经下载下来的列表页html保存到的目录（文件夹）
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            //输出excel表格包含的列
            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                "detailPageUrl",
                "detailPageName", 
                "cookie",
                "grabStatus", 
                "giveUpGrab",
                "date",
                "title",
                "pageIndex"}); 
            
            //输出文件的本地路径，此输出文件是项目“节能搜索低碳获取详情页”的输入文件
            string resultFilePath = Path.Combine(exportDir, "节能搜索低碳获取详情页.xlsx");

            //输出对象
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic);

            //循环输入文件中的所有行
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                string pageUrl = listSheet.PageUrlList[i];
                Dictionary<string, string> row = listSheet.GetRow(i);

                //输入文件中，此行的pageIndex列的值
                string pageIndex = row["pageIndex"];

                //已经下载下来的html的本地保存地址
                string localFilePath = this.RunPage.GetFilePath(pageUrl, pageSourceDir);
                TextReader tr = null;

                try
                { 
                    //读取已经获取到本地的列表页html，并加载到HtmlDocument对象中（系统提供了构造HtmlDocument对象的方法，写这段代码的时候忘记用了）                       
                    tr = new StreamReader(localFilePath);
                    string webPageHtml = tr.ReadToEnd();
                    HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                    htmlDoc.LoadHtml(webPageHtml);

                    //利用xpath获取列表页中包含的详情页url、发布标题title、发布时间
                    HtmlNodeCollection allItemParentNodes = htmlDoc.DocumentNode.SelectNodes("//div[@class=\"c wrap\"]/ul[@class=\"wrap\"]/li");
                    if (allItemParentNodes != null)
                    {
                        foreach (HtmlNode itemParentNode in allItemParentNodes)
                        {
                            HtmlNode itemNode = itemParentNode.SelectSingleNode("./div[1]/h5/a");
                            string url = itemNode.Attributes["href"].Value;
                            string title = CommonUtil.HtmlDecode(itemNode.InnerText).Trim();

                            HtmlNode dateNode = itemParentNode.SelectSingleNode("./div[@class=\"adds\"]");
                            string dateStr = CommonUtil.HtmlDecode(dateNode.InnerText).Trim().Substring(5, 10);
                            DateTime date = DateTime.ParseExact(dateStr, "yyyy-MM-dd", System.Globalization.CultureInfo.CurrentCulture);

                            //只记录某个日期至今的详情页url及其他属性
                            if (date >= toDate)
                            {
                                Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                f2vs.Add("detailPageUrl", url);
                                f2vs.Add("detailPageName", url);
                                f2vs.Add("title", title);
                                f2vs.Add("date", dateStr);
                                f2vs.Add("pageIndex", pageIndex);
                                resultEW.AddRow(f2vs);
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

            //保存到硬盘
            resultEW.SaveToDisk();

            return true;
        }
        #endregion 
    }
}