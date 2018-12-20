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
using NetDataAccess.Base.Reader;

namespace NetDataAccess.Extended.GuoXueDaShi.LiShiShiJian
{
    public class GetLiShiShiJianQueryPage : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        { 
            this.GetYearPageUrls(listSheet);
            return true;
        }

        private void GetYearPageUrls(IListSheet listSheet)
        {
            string sourceDir = this.RunPage.GetDetailSourceFileDir();
            ExcelWriter resultEW = this.CreateResultWriter();
            Dictionary<string, bool> pageUrlDic = new Dictionary<string, bool>();
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> listRow = listSheet.GetRow(i);
                string yearValue = listRow["yearValue"];
                HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);

                HtmlNodeCollection yearPageUrlNodes = htmlDoc.DocumentNode.SelectNodes("//div[@class=\"info_content zj clearfix\"]/dl/dd/a");
                if (yearPageUrlNodes == null || yearPageUrlNodes.Count == 0)
                {
                    string errorInfo = "未查询到" + yearValue + "年的事件页面";
                    this.RunPage.InvokeAppendLogText(errorInfo, LogLevelType.Error, true);
                }
                else if (yearPageUrlNodes.Count > 1)
                {
                    throw new Exception("查询到多余一个的" + yearValue + "年的事件页面");
                }
                else
                {
                    HtmlNode yearPageUrlNode = yearPageUrlNodes[0];
                    string pageUrl = yearPageUrlNode.GetAttributeValue("href", "");
                    string fullPageUrl = "http://www.guoxuedashi.com" + pageUrl;
                    if (!pageUrlDic.ContainsKey(fullPageUrl))
                    {
                        pageUrlDic.Add(fullPageUrl, true);
                        string yearName = CommonUtil.HtmlDecode(yearPageUrlNode.InnerText).Trim(); 
                        Dictionary<string, string> resultRow = new Dictionary<string, string>();
                        resultRow.Add("detailPageUrl", fullPageUrl);
                        resultRow.Add("detailPageName", yearValue);
                        resultRow.Add("yearName", yearName);
                        resultRow.Add("yearValue", yearValue);
                        resultEW.AddRow(resultRow);
                    }
                }
            }

            resultEW.SaveToDisk();
        }

        private ExcelWriter CreateResultWriter()
        {
            String exportDir = this.RunPage.GetExportDir();
            string resultFilePath = Path.Combine(exportDir, "国学大师_历史事件_年份页面.xlsx");

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("yearName", 5);
            resultColumnDic.Add("yearValue", 6); 
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }
    }
}