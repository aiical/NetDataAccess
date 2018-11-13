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

namespace NetDataAccess.Extended.Keywords.WanFang
{
    public class GetKeywordSearchAllPages : ExternalRunWebPage
    { 
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            try
            {
                string[] parameters = this.Parameters.Split(new string[] { }, StringSplitOptions.RemoveEmptyEntries);
                string qiKanListFilePath = parameters[0];

                Dictionary<string, bool> qiKanDic = this.GetAllQiKans(qiKanListFilePath);

                this.GetAllPages(listSheet, qiKanDic);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return true;
        }

        private ExcelWriter GetPageUrlsExcelWriter(int fileIndex)
        {
            String exportDir = this.RunPage.GetExportDir();
            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[] { 
                "keyword", 
                "品类", 
                "词类型",
                "keywords" });

            string resultFilePath = Path.Combine(exportDir, "万方期刊_专业关键词_关键词列表_" + fileIndex.ToString() + ".xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic);
            return resultEW;
        }

        private ExcelWriter GetPinLeiKeywordExcelWriter()
        {
            String exportDir = this.RunPage.GetExportDir();
            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[] {  
                "品类", 
                "词汇",
                "出现次数"});

            string resultFilePath = Path.Combine(exportDir, "万方期刊_品类_行业词汇.xlsx");
            Dictionary<string, string> columnFormats = new Dictionary<string, string>();
            columnFormats.Add("出现次数", "#0");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, columnFormats);
            return resultEW;
        }

        private Dictionary<string, bool> GetAllQiKans(string filePath)
        {
            Dictionary<string, bool> qiKanDic = new Dictionary<string, bool>();
            ExcelReader er = new ExcelReader(filePath);
            int rowCount = er.GetRowCount();
            for (int i = 0; i < rowCount; i++)
            {
                Dictionary<string, string> row = er.GetFieldValues(i);
                string qiKan = row["perio_title"];
                qiKanDic.Add(qiKan, true);
            } 
            return qiKanDic;
        }

        private void AddToPinLeiToKeywordList(Dictionary<string, Dictionary<string,int>> pinLeiToKeywordList, string pinLei, string keywordStr)
        {
            if (!pinLeiToKeywordList.ContainsKey(pinLei))
            {
                pinLeiToKeywordList.Add(pinLei, new Dictionary<string, int>());
            }
            Dictionary<string, int> keywordList = pinLeiToKeywordList[pinLei];
            string[] keywords = keywordStr.Split(new string[] { "   " }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < keywords.Length; i++)
            {
                string keyword = keywords[i].Trim();
                if (!keywordList.ContainsKey(keyword))
                {
                    keywordList.Add(keyword, 1);
                }
                else
                {
                    keywordList[keyword] = keywordList[keyword] + 1;
                }
            }
        }

        private void GetAllPages(IListSheet listSheet, Dictionary<string, bool> qiKanDic)
        {
            Dictionary<string, Dictionary<string, int>> pinLeiToKeywordList = new Dictionary<string, Dictionary<string, int>>();

            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            int fileIndex = 1;
            ExcelWriter ew = this.GetPageUrlsExcelWriter(fileIndex);
            Dictionary<string, string> idDic = new Dictionary<string, string>();
            int rowCount = 0;
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                if (rowCount % 1000 == 0)
                {
                    this.RunPage.InvokeAppendLogText("已处理到: fileIndex = " + fileIndex.ToString() + ", rowCount = " + rowCount.ToString(), LogLevelType.System, true);
                }

                if (rowCount >= 500000)
                {
                    if (ew != null)
                    {
                        ew.SaveToDisk();
                    }
                    fileIndex++;
                    ew = this.GetPageUrlsExcelWriter(fileIndex);
                    rowCount = 0;
                }

                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"];

                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string localFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir);

                    try
                    {
                        string pageFileText = FileHelper.GetTextFromFile(localFilePath);
                        HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                        htmlDoc.LoadHtml(pageFileText);
                        HtmlNodeCollection keywordNodes = htmlDoc.DocumentNode.SelectNodes("//ul[@class=\"list_ul\"]/li[@class=\"zi\"]/p[@class=\"greencolor\"]");
                        if(keywordNodes!=null){
                            foreach (HtmlNode keywordNode in keywordNodes)
                            {
                                HtmlNodeCollection linkNodes = keywordNode.ParentNode.ParentNode.SelectNodes("./li[@class=\"greencolor\"]/a");
                                bool isCorrectQiKan = false;
                                if (linkNodes != null)
                                {
                                    foreach (HtmlNode linkNode in linkNodes)
                                    {
                                        string text = CommonUtil.HtmlDecode(linkNode.InnerText).Trim();
                                        if (text.StartsWith("《") && text.EndsWith("》"))
                                        {
                                            string qiKan = text.Substring(1, text.Length - 2).Trim();
                                            if (qiKanDic.ContainsKey(qiKan))
                                            {
                                                isCorrectQiKan = true;
                                                break;
                                            }
                                        }
                                    }
                                }

                                if (isCorrectQiKan)
                                {
                                    string keywordStr = CommonUtil.HtmlDecode(keywordNode.InnerText).Replace("关键词：", "").Trim();
                                    string pinLei= row["品类"];

                                    Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                    f2vs.Add("keyword", row["keyword"]);
                                    f2vs.Add("品类",pinLei );
                                    f2vs.Add("词类型", row["词类型"]);
                                    f2vs.Add("keywords", keywordStr);
                                    ew.AddRow(f2vs);
                                    rowCount++;

                                    AddToPinLeiToKeywordList(pinLeiToKeywordList, pinLei, keywordStr);
                                }
                            }
                        } 
                    }
                    catch (Exception ex)
                    {
                        this.RunPage.InvokeAppendLogText(ex.Message + ". detailUrl = " + detailUrl, LogLevelType.Error, true);
                        throw ex;

                    }
                }
            }
            ew.SaveToDisk();

            ExcelWriter pinLeiKeywordExcelWriter = this.GetPinLeiKeywordExcelWriter();
            foreach (string pinLei in pinLeiToKeywordList.Keys)
            {
                Dictionary<string, int> keywordList = pinLeiToKeywordList[pinLei];
                foreach (string keyword in keywordList.Keys)
                {
                    int count = keywordList[keyword];

                    Dictionary<string, object> row = new Dictionary<string, object>();
                    row.Add("品类", pinLei);
                    row.Add("词汇", keyword);
                    row.Add("出现次数", count);
                    pinLeiKeywordExcelWriter.AddRow(row);
                }
            }
            pinLeiKeywordExcelWriter.SaveToDisk();
        } 
    }
}