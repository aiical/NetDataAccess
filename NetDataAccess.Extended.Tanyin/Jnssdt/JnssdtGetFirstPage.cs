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
    /// 节能搜索低碳获取首页
    /// 运行此程序前，系统已经提前爬取了listSheet中指定的节能搜索低碳获取首页html
    /// 利用此程序，获取到节能搜索中搜索低碳可以查询到多少页记录，并按照url规则形成所有页面的url，为“节能搜索低碳获取列表页”提供下载列表
    /// </summary>
    public class JnssdtGetFirstPage : CustomProgramBase
    {
        #region 入口函数
        /// <summary>
        /// 入口函数
        /// </summary>
        /// <param name="parameters">“扩展程序配置”信息中的parameters属性值</param>
        /// <param name="listSheet">输入文件，记录了要下载的所有页面，本项目只含一条记录</param>
        /// <returns></returns>
        public bool Run(string parameters, IListSheet listSheet)
        {
            return GetAllListPageUrl(parameters, listSheet);
        }
        #endregion

        #region 获取所有列表页地址
        private bool GetAllListPageUrl(string parameters, IListSheet listSheet)
        { 
            //已经下载下来的首页html保存到的目录（文件夹）
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            
            //输出excel表格包含的列
            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                "detailPageUrl",
                "detailPageName", 
                "cookie",
                "grabStatus", 
                "giveUpGrab",
                "pageIndex"});

            //输出目录（从配置中获取）
            string exportDir = this.RunPage.GetExportDir();

            //输出文件的本地路径，此输出文件是项目“节能搜索低碳获取列表页”的输入文件
            string resultFilePath = Path.Combine(exportDir, "节能搜索低碳获取列表页.xlsx");

            //输出对象
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic);

            //从下载到的首页html中，获取列表页个数，并形成所有列表页url
            GetAllListPageUrl(listSheet, pageSourceDir, resultEW);

            //保存到硬盘
            resultEW.SaveToDisk();

            return true;
        }
        #endregion

        #region 获取所有列表页地址
        /// <summary>
        /// 从下载到的首页html中，获取列表页个数，并形成所有列表页url
        /// </summary>
        /// <param name="listSheet"></param>
        /// <param name="pageSourceDir"></param>
        /// <param name="resultEW"></param>
        private void GetAllListPageUrl(IListSheet listSheet, string pageSourceDir, ExcelWriter resultEW)
        {
            //listSheet中只有一条记录
            string pageUrl = listSheet.PageUrlList[0];
            Dictionary<string, string> row = listSheet.GetRow(0); 
            string localFilePath = this.RunPage.GetFilePath(pageUrl, pageSourceDir);
            TextReader tr = null;

            try
            {
                tr = new StreamReader(localFilePath);
                string webPageHtml = tr.ReadToEnd();
                HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                htmlDoc.LoadHtml(webPageHtml);

                //获取导航栏元素
                HtmlNodeCollection allListPageLinkNodes = htmlDoc.DocumentNode.SelectNodes("//div[@id=\"pages\"]/a");
                if (allListPageLinkNodes != null)
                {
                    //获取显示最大页码数的那个元素
                    HtmlNode lastPageLinkNode = allListPageLinkNodes[allListPageLinkNodes.Count - 3];
                    string lastPageNumStr = CommonUtil.HtmlDecode(lastPageLinkNode.InnerText).Trim();
                    int lastPageNum = int.Parse(lastPageNumStr);

                    //构造列表页url并保存
                    for (int i = 0; i < lastPageNum; i++)
                    {
                        int pageIndex = i + 1;
                        string url = "http://so.ces.cn/index.php?typeid=127&q=%E4%BD%8E%E7%A2%B3&page=" + pageIndex.ToString();
                        Dictionary<string, string> f2vs = new Dictionary<string, string>();
                        f2vs.Add("detailPageUrl", url);
                        f2vs.Add("detailPageName", pageIndex.ToString());
                        f2vs.Add("pageIndex", pageIndex.ToString());
                        resultEW.AddRow(f2vs);
                    }
                }
            }
            catch (Exception ex)
            {
                this.RunPage.InvokeAppendLogText("读取出错. url = " + pageUrl + ". " + ex.Message, LogLevelType.Error, true);
                throw ex;
            }
        }
        #endregion
    }
}