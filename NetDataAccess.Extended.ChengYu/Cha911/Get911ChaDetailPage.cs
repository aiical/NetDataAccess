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

namespace NetDataAccess.Extended.ChengYu.Cha911
{
    public class Get911ChaDetailPage : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        { 
            this.GetDetailInfos(listSheet);
            return true;
        }

        private string GetProcessedHtml(string fileHtml)
        { 
            int fileHtmlLength = fileHtml.Length;
            StringBuilder htmlStrBuilder = new StringBuilder();
            bool inBookName = false;
            for (int i = 0; i < fileHtmlLength; i++)
            {
                bool needAddToNewString = true;
                char oneLetter = fileHtml[i];
                if (oneLetter == '《')
                {
                    inBookName = true;
                }
                else if (oneLetter == '》')
                {
                    if (inBookName)
                    {
                        inBookName = false;
                    }
                }
                else if (oneLetter == '<')
                {
                    var nextLetter = fileHtml[i + 1];
                    if (CommonUtil.CheckStringChineseReg(nextLetter.ToString()))
                    {
                        needAddToNewString = false;
                    }
                }
                else if (oneLetter == '>')
                {
                    var preLetter = fileHtml[i - 1];
                    if (CommonUtil.CheckStringChineseReg(preLetter.ToString()))
                    {
                        needAddToNewString = false;
                    }
                }
                if (needAddToNewString)
                {
                    htmlStrBuilder.Append(oneLetter);
                }
            }
            return htmlStrBuilder.ToString();
        }

        private void GetDetailInfos(IListSheet listSheet)
        {
            string sourceDir = this.RunPage.GetDetailSourceFileDir();
            ExcelWriter resultEW = this.CreateResultWriter();
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> listRow = listSheet.GetRow(i);
                string detailPageUrl = listRow[SysConfig.DetailPageUrlFieldName];
                string filePath = this.RunPage.GetFilePath(detailPageUrl, sourceDir);
                string fileHtml = FileHelper.GetTextFromFile(filePath);
                try
                {
                    fileHtml = this.GetProcessedHtml(fileHtml);

                    HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                    htmlDoc.LoadHtml(fileHtml);
                    HtmlNode nameNode = htmlDoc.DocumentNode.SelectSingleNode("//div[@class=\"mcon\"]/h2");
                    string name = CommonUtil.HtmlDecode(nameNode.InnerText).Trim();

                    HtmlNode pinyinNode = htmlDoc.DocumentNode.SelectSingleNode("//div[@class=\"mcon\"]/div[@class=\"gray mt f16\"]");
                    string pinyin = CommonUtil.HtmlDecode(pinyinNode.InnerText).Trim();

                    Dictionary<string, string> resultRow = new Dictionary<string, string>();
                    resultRow.Add("成语", name);
                    resultRow.Add("拼音", pinyin);

                    HtmlNodeCollection propertyNodes = htmlDoc.DocumentNode.SelectNodes("//div[@class=\"mcon bt noi f14\"]/p");
                    if (propertyNodes == null)
                    {
                        throw new Exception("获取成语属性失败. name = " + name + ".");
                    }
                    else
                    {
                        foreach (HtmlNode propertyNode in propertyNodes)
                        {
                            try
                            {
                                //this.RunPage.InvokeAppendLogText(propertyNode.InnerText, LogLevelType.System, true);
                                HtmlNode propertyNameNode = propertyNode.SelectSingleNode("./span");
                                if (propertyNameNode != null)
                                {
                                    string propertyName = CommonUtil.HtmlDecode(propertyNameNode.InnerText).Trim();
                                    StringBuilder valueStrBuilder = new StringBuilder();
                                    HtmlNode valueNode = propertyNameNode.NextSibling;
                                    while (valueNode != null)
                                    {
                                        string partValue = CommonUtil.HtmlDecode(valueNode.InnerText).Replace("\r\n", "").Replace("911cha.com", "");
                                        valueStrBuilder.Append(partValue);
                                        valueNode = valueNode.NextSibling;
                                    }
                                    resultRow.Add(propertyName, valueStrBuilder.ToString());
                                }
                            }
                            catch (Exception ex)
                            {
                                throw new Exception("解析成语属性出错. propertyNode = " + propertyNode.OuterHtml + "." + ex.Message);
                            }
                        }
                    }
                    resultEW.AddRow(resultRow);
                }
                catch (Exception ex)
                {
                    throw new Exception("获取成语信息出错. filePath = " + filePath + "." + ex.Message);
                }
            }

            resultEW.SaveToDisk();
        }

        private ExcelWriter CreateResultWriter()
        {
            String exportDir = this.RunPage.GetExportDir();
            string resultFilePath = Path.Combine(exportDir, "成语_911Cha_详细信息.xlsx");

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("成语", 0);
            resultColumnDic.Add("拼音", 1);
            resultColumnDic.Add("成语解释", 2);
            resultColumnDic.Add("成语出处", 3);
            resultColumnDic.Add("成语繁体", 4);
            resultColumnDic.Add("成语简拼", 5);
            resultColumnDic.Add("成语注音", 6);
            resultColumnDic.Add("常用程度", 7);
            resultColumnDic.Add("成语字数", 8);
            resultColumnDic.Add("感情色彩", 9);
            resultColumnDic.Add("成语用法", 10);
            resultColumnDic.Add("成语结构", 11);
            resultColumnDic.Add("成语年代", 12);
            resultColumnDic.Add("成语例子", 13);
            resultColumnDic.Add("英语翻译", 14);
            resultColumnDic.Add("俄语翻译", 15);
            resultColumnDic.Add("其他翻译", 16);
            resultColumnDic.Add("成语谜语", 17);
            resultColumnDic.Add("成语接龙", 18);
            resultColumnDic.Add("近义词", 19);
            resultColumnDic.Add("反义词", 20);
            resultColumnDic.Add("成语辨析", 21);
            resultColumnDic.Add("成语辨形", 22);
            resultColumnDic.Add("成语正音", 23);
            resultColumnDic.Add("成语故事", 24);
            resultColumnDic.Add("日语翻译", 25);
            resultColumnDic.Add("成语歇后语", 26); 
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }
    }
}