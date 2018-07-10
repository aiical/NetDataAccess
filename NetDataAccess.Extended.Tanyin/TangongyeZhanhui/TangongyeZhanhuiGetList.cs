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
    /// Tangongye展会获取列表页
    /// 不使用工具自带的自动抓取功能。
    /// 在扩展程序类TangongyeZhanhuiGetList中调用系统提供的API抓取列表页，然后分析页面获取详情页地址
    /// </summary>
    public class TangongyeZhanhuiGetList : CustomProgramBase
    {
        #region 入口函数
        /// <summary>
        /// 入口函数
        /// </summary>
        /// <param name="parameters">“扩展程序配置”信息中的parameters属性值</param>
        /// <param name="listSheet">输入文件，记录了要下载的所有页面，本项目未使用此文件</param>
        /// <returns></returns>
        public bool Run(string parameters, IListSheet listSheet)
        {
            return GetAllListPageUrl(parameters);
        }
        #endregion

        #region 获取所有列表页
        /// <summary>
        /// 获取所有列表页
        /// </summary>
        /// <param name="parameters"></param>
        /// <returns></returns>
        private bool GetAllListPageUrl(string parameters)
        {
            try
            {
                //配置文件中，扩展程序指定的parameters属性值
                string[] parameterArray = parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);

                //参数中指定的输出文件夹
                string exportDir = parameterArray[0];

                //需要抓取此日期后的展会信息
                DateTime toDate = DateTime.ParseExact(parameterArray[1], "yyyy-MM-dd", System.Globalization.CultureInfo.CurrentCulture);
                DateTime todayDate = DateTime.Now.Date;

                //下载的列表页存放的位置
                string pageSourceDir = Path.Combine(exportDir, "Detail");

                //输出所有详情页地址及相关属性
                Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                "detailPageUrl",
                "detailPageName", 
                "cookie",
                "grabStatus", 
                "giveUpGrab",
                "title",
                "date",
                "pageNum"});

                //输出文件的本地路径，此输出文件是项目“Tangongye展会获取详情页”的输入文件
                string resultFilePath = Path.Combine(exportDir, "Export\\Tangongye展会获取详情页.xlsx");

                //初始化输出文件对象，以excel文件格式输出
                ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
                 
                //是否可以继续抓取，循环中用到
                bool canGoon = true;

                //记录当前页码，递增
                int pageIndex = 1;

                //指定获取到的网页编码方式
                Encoding encoding = Encoding.UTF8;

                //记录解析获得的详情页地址，用来去重
                Dictionary<string, string> allPageUrls = new Dictionary<string, string>();

                //循环获取符合条件（日期限制）的详情页地址
                while (canGoon)
                {
                    //将当前进度信息展示在界面中
                    this.RunPage.InvokeAppendLogText("正在获取第" + pageIndex.ToString() + "页...", LogLevelType.System, true);
                    
                    //默认没有在此次获取的列表页中找到符合条件的详情页，即默认会跳出循环
                    canGoon = false;

                    //本次要获取的列表页地址
                    string listPageUrl = "http://www.tangongye.com/news/NewsList.aspx?page=" + pageIndex.ToString() + "&categoryId=22";

                    //通过工具提供的api获取这个列表页
                    string webPageHtml = "";// this.RunPage.GetTextByRequest(listPageUrl, false, 1, 30, encoding, null, null, false, Proj_DataAccessType.WebRequestHtml, null);

                    //获取到的列表页保存到的本地路径
                    string localSourcePageFilePath = this.RunPage.GetFilePath(listPageUrl, pageSourceDir);

                    //保存到本地
                    this.RunPage.SaveFile(webPageHtml, localSourcePageFilePath, encoding);

                    //使用HtmlDocument加载获取到的列表页html
                    HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                    htmlDoc.LoadHtml(webPageHtml);

                    //解析获取所有列表页中的详情页条目
                    HtmlNodeCollection allPageNodes = htmlDoc.DocumentNode.SelectNodes("//div[@id=\"news-list\"]/ul/li");
                    if (allPageNodes != null)
                    {
                        foreach (HtmlNode pageNode in allPageNodes)
                        {
                            //获取详情页发布日期
                            HtmlNode dateNode = pageNode.SelectSingleNode("./span");
                            string dateStr = CommonUtil.HtmlDecode(dateNode.InnerText.Trim()).Trim();
                            DateTime pageDate = DateTime.ParseExact(dateStr, "yyyy-MM-dd", System.Globalization.CultureInfo.CurrentCulture);

                            //过滤出符合条件的详情页
                            if (pageDate >= toDate)
                            {
                                HtmlNode linkNode = pageNode.SelectSingleNode("./a");
                                string linkUrl = linkNode.Attributes["href"].Value;
                                if (!allPageUrls.ContainsKey(linkUrl))
                                {
                                    allPageUrls.Add(linkUrl, null);
                                    string title = CommonUtil.HtmlDecode(linkNode.InnerText.Trim()).Trim();
                                    string detailPageName = linkUrl.Substring(linkUrl.LastIndexOf("=") + 1);
                                    string detailPageUrl = "http://www.tangongye.com/news/" + linkUrl;

                                    //记录在内存里，待保存到本地
                                    Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                    f2vs.Add("detailPageUrl", detailPageUrl);
                                    f2vs.Add("detailPageName", detailPageName);
                                    f2vs.Add("title", title);
                                    f2vs.Add("pageNum", pageIndex.ToString());
                                    f2vs.Add("date", dateStr);
                                    resultEW.AddRow(f2vs);
                                    canGoon = true;
                                }
                            }
                        }
                    }
                    pageIndex++;
                }

                //保存到本地
                resultEW.SaveToDisk();
            }
            catch (Exception ex)
            {
                throw new Exception("读取列表页失败!", ex);
            }

            return true;
        }
        #endregion
    }
}