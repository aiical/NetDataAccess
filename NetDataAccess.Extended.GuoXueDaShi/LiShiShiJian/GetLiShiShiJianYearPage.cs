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
using Microsoft.International.Converters.TraditionalChineseToSimplifiedConverter;

namespace NetDataAccess.Extended.GuoXueDaShi.LiShiShiJian
{
    public class GetLiShiShiJianYearPage : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        { 
            this.GetYearPageInfo(listSheet);
            return true;
        }

        private void GetYearPageInfo(IListSheet listSheet)
        {
            string sourceDir = this.RunPage.GetDetailSourceFileDir();
            ExcelWriter resultEW = this.CreateResultWriter(); 
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> listRow = listSheet.GetRow(i);
                string yearValue = listRow["yearValue"];
                string yearName = listRow["yearName"];
                string pageUrl = listRow[SysConfig.DetailPageUrlFieldName];
                HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);

                HtmlNodeCollection pNodes = htmlDoc.DocumentNode.SelectNodes("//span[@class=\"table2\"]/p");
                StringBuilder textInfo = new StringBuilder();
                if (pNodes == null || pNodes.Count == 0)
                {
                    string errorInfo = "未查询到" + yearValue + "年的事件内容";
                    this.RunPage.InvokeAppendLogText(errorInfo, LogLevelType.Error, true);
                }
                else
                {
                    bool startText = false;
                    foreach (HtmlNode pNode in pNodes)
                    {
                        string pText = CommonUtil.HtmlDecode(pNode.InnerText).Trim();
                        if (startText)
                        {
                            textInfo.AppendLine(pText);
                        }
                        else if (pText == "大事記")
                        {
                            startText = true;
                        }
                    }

                    string eventText = textInfo.ToString();
                    string eventTextSimple1 = CharProcessor.ConverterChinese(eventText, ChineseConversionDirection.TraditionalToSimplified);
                    string eventTextSimple2 = Microsoft.VisualBasic.Strings.StrConv(eventText, Microsoft.VisualBasic.VbStrConv.SimplifiedChinese, 0);
                    Dictionary<string, string> resultRow = new Dictionary<string, string>();
                    resultRow.Add("url", pageUrl);
                    resultRow.Add("yearName", yearName);
                    resultRow.Add("yearValue", yearValue);
                    resultRow.Add("eventText", eventText);
                    resultRow.Add("eventTextSimple1", eventTextSimple1);
                    resultRow.Add("eventTextSimple2", eventTextSimple2);
                    resultEW.AddRow(resultRow);
                }
            }

            resultEW.SaveToDisk();
        }

        private ExcelWriter CreateResultWriter()
        {
            String exportDir = this.RunPage.GetExportDir();
            string resultFilePath = Path.Combine(exportDir, "国学大师_历史事件_年份事件.xlsx");

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("yearName", 0);
            resultColumnDic.Add("yearValue", 1);
            resultColumnDic.Add("url", 2);
            resultColumnDic.Add("eventText", 3);
            resultColumnDic.Add("eventTextSimple1", 4);
            resultColumnDic.Add("eventTextSimple2", 5);
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }
    }
}