using HtmlAgilityPack;
using NetDataAccess.Base.Browser;
using NetDataAccess.Base.Common;
using NetDataAccess.Base.Config;
using NetDataAccess.Base.DB;
using NetDataAccess.Base.EnumTypes;
using NetDataAccess.Base.Reader;
using NetDataAccess.Base.UI;
using NetDataAccess.Base.Web;
using NetDataAccess.Base.Writer;
using NetDataAccess.Extended.LunWen.Common;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace NetDataAccess.Extended.LunWen.ScienceDirect
{ 
    public class GetAllLunWenPdfUrls : ExternalRunWebPage
    {
        private CookieContainer _CC = null;


        public override void WebRequestHtml_BeforeSendRequest(string pageUrl, Dictionary<string, string> listRow, NDAWebClient client)
        {
            this._CC = new CookieContainer();
            client.CookieContainer = this._CC;
        } 

        public override void CheckRequestCompleteFile(string webPageText, Dictionary<string, string> listRow)
        {
            if (!webPageText.Contains("redirect-message"))
            {
                throw new BlockedException("被封了");
            }
            else
            {
                string pageUrl = listRow[SysConfig.DetailPageUrlFieldName];
                HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                htmlDoc.LoadHtml(webPageText);


                HtmlNode linkNode = htmlDoc.DocumentNode.SelectSingleNode("//div[@id=\"redirect-message\"]/p/a");
                string pdfUrl = linkNode.GetAttributeValue("href", "");


                string sourceDir = this.RunPage.GetDetailSourceFileDir();
                string pdfFilePath = this.RunPage.GetFilePath(pdfUrl, sourceDir);
                try
                {
                    NDAWebClient wc = new NDAWebClient();
                    wc.CookieContainer = this._CC;
                    //string cookie = this._CC.GetCookieHeader(new Uri(pageUrl));
                    byte[] bytes = wc.DownloadData(pdfUrl);
                    FileStream fs = null;
                    try
                    {
                        fs = new FileStream(pdfFilePath, FileMode.Create, FileAccess.Write);
                        fs.Write(bytes, 0, bytes.Length);
                        fs.Flush();
                    }
                    catch (Exception ex)
                    {
                        this.RunPage.InvokeAppendLogText("保存pdf文件出错", LogLevelType.Error, true);
                        throw ex;
                    }
                    finally
                    {
                        if (fs != null)
                        {
                            fs.Close();
                            fs.Dispose();
                            fs = null;
                        }
                    }
                }
                catch (Exception ex)
                {
                    this.RunPage.InvokeAppendLogText("获取pdf文件出错", LogLevelType.Error, true);
                    throw ex;
                }
            }
        }

        public override bool AfterGrabOneCatchException(string pageUrl, Dictionary<string, string> listRow, Exception ex)
        {
            if (ex is BlockedException)
            {
                this.RunPage.InvokeAppendLogText("封号了，休息10分钟", LogLevelType.Error, true);
                Thread.Sleep(1000 * 60 * 10);
            }
            return false;
        }

        public override bool AfterAllGrab(IListSheet listSheet)
        {
            this.ConvertToTxt(listSheet);
            this.FindKeywords(listSheet);
            return true;
        }
        private void ConvertToTxt(IListSheet listSheet)
        {
            try
            {
                string sourceDir = this.RunPage.GetDetailSourceFileDir();
                string exportDir = this.RunPage.GetExportDir();
                string pdfUrlFilePath = Path.Combine(exportDir, "论文_ScienceDirect_论文PDF页.xlsx");
                ExcelWriter pdfUrlWriter = this.GetDownloadPdfExcelWriter(pdfUrlFilePath);

                for (int i = 0; i < listSheet.RowCount; i++)
                {
                    this.RunPage.InvokeAppendLogText("已转换" + i.ToString() + "/" + listSheet.RowCount.ToString(), LogLevelType.System, true);

                    Dictionary<string, string> listRow = listSheet.GetRow(i);
                    string pageUrl = listRow[SysConfig.DetailPageUrlFieldName];

                    bool giveUp = "Y".Equals(listRow[SysConfig.GiveUpGrabFieldName]);
                    if (!giveUp)
                    {
                        try
                        {
                            string textFileDir = this.RunPage.GetReadFilePath(pageUrl, exportDir);
                            string fullTextFilePath = Path.Combine(textFileDir, "allText.txt");
                            if (!File.Exists(fullTextFilePath))
                            {
                                HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                                HtmlNode linkNode = htmlDoc.DocumentNode.SelectSingleNode("//div[@id=\"redirect-message\"]/p/a");
                                string pdfUrl = linkNode.GetAttributeValue("href", "");
                                string pdfFilePath = this.RunPage.GetFilePath(pdfUrl, sourceDir);
                                if (!Directory.Exists(textFileDir))
                                {
                                    Directory.CreateDirectory(textFileDir);
                                }
                                string[] pdfPartFilePaths = PdfSpliter.ExtractPages(pdfFilePath, textFileDir);
                                StringBuilder fullText = new StringBuilder();
                                for (int j = 0; j < pdfPartFilePaths.Length; j++)
                                {
                                    string pdfPartFilePath = pdfPartFilePaths[j];
                                    string textPartFilePath = Path.Combine(textFileDir, (j + 1).ToString() + ".txt");
                                    try
                                    {
                                        Pdf2Txt.Pdf2TxtByITextSharp(pdfPartFilePath, textPartFilePath, true);

                                        string text = FileHelper.GetTextFromFile(textPartFilePath, Encoding.UTF8);
                                        fullText.Append(text);
                                    }
                                    catch (Exception pdf2TxtEx)
                                    {
                                        if (pdf2TxtEx.Message.Contains("System.FormatException"))
                                        {
                                            this.RunPage.InvokeAppendLogText("转换txt失败, pdfPartFilePath = " + pdfPartFilePath, LogLevelType.Error, true);
                                        }
                                        else
                                        {
                                            throw pdf2TxtEx;
                                        }
                                    }
                                }
                                FileHelper.SaveTextToFile(fullText.ToString(), fullTextFilePath, Encoding.UTF8);
                            }

                            Dictionary<string, string> pdfUrlRow = new Dictionary<string, string>();
                            pdfUrlRow.Add("publication", listRow["publication"]);
                            pdfUrlRow.Add("host", listRow["host"]);
                            pdfUrlRow.Add("title", listRow["title"]);
                            pdfUrlRow.Add("authors", listRow["authors"]);
                            pdfUrlRow.Add("abstracts", listRow["abstracts"]);
                            pdfUrlRow.Add("refs", listRow["refs"]);
                            pdfUrlRow.Add("pageUrl", pageUrl);
                            pdfUrlRow.Add("txtUrl", fullTextFilePath);
                            pdfUrlWriter.AddRow(pdfUrlRow);
                        }
                        catch (Exception ex)
                        {
                            string filePath = this.RunPage.GetFilePath(pageUrl, sourceDir);
                            this.RunPage.InvokeAppendLogText("错误，filePath = " + filePath + ", pageUrl = " + pageUrl, LogLevelType.Error, true);
                            throw ex;
                        }
                    }
                }

                pdfUrlWriter.SaveToDisk();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        private void FindKeywords(IListSheet listSheet)
        {
            try
            {
                string sourceDir = this.RunPage.GetDetailSourceFileDir();
                string exportDir = this.RunPage.GetExportDir();
                string keywordsFilePath = Path.Combine(exportDir, "论文_ScienceDirect_论文关键词.xlsx");
                ExcelWriter keywordsWriter = this.GetKeywordsExcelWriter(keywordsFilePath);

                string allFileKeywordsFilePath = Path.Combine(exportDir, "论文_ScienceDirect_论文关键词_所有文章.xlsx");
                ExcelWriter allFileKeywordsWriter = this.GetAllFileKeywordsExcelWriter(allFileKeywordsFilePath);

                for (int i = 0; i < listSheet.RowCount; i++)
                {
                    this.RunPage.InvokeAppendLogText("已提取Keywords" + i.ToString() + "/" + listSheet.RowCount.ToString(), LogLevelType.System, true);

                    Dictionary<string, string> listRow = listSheet.GetRow(i);
                    string pageUrl = listRow[SysConfig.DetailPageUrlFieldName];

                    bool giveUp = "Y".Equals(listRow[SysConfig.GiveUpGrabFieldName]);
                    if (!giveUp)
                    {
                        int year = this.GetYear(listRow["host"]);
                        string textFileDir = this.RunPage.GetReadFilePath(pageUrl, exportDir);
                        string keywordsTextFilePath = Path.Combine(textFileDir, "keywords.txt");

                        try
                        {
                            string[] keywordArray = null;
                            if (!File.Exists(keywordsTextFilePath))
                            {
                                string allTextFilePath = Path.Combine(textFileDir, "allText.txt");
                                keywordArray = this.FindKeywords(allTextFilePath);
                                string keywordStr = CommonUtil.StringArrayToString(keywordArray, "\r\n");
                                FileHelper.SaveTextToFile(keywordStr, keywordsTextFilePath);
                            }
                            else
                            {
                                string keywordsStr = FileHelper.GetTextFromFile(keywordsTextFilePath);
                                keywordArray = keywordsStr.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
                            }

                            if (keywordArray != null && keywordArray.Length > 0)
                            {
                                foreach (string keyword in keywordArray)
                                {
                                    Dictionary<string, object> keywordRow = new Dictionary<string, object>();
                                    keywordRow.Add("publication", listRow["publication"]);
                                    keywordRow.Add("host", listRow["host"]);
                                    keywordRow.Add("title", listRow["title"]);
                                    keywordRow.Add("authors", listRow["authors"]);
                                    keywordRow.Add("year", year);
                                    keywordRow.Add("pageUrl", pageUrl);
                                    keywordRow.Add("keyword", keyword);
                                    keywordsWriter.AddRow(keywordRow);
                                }
                            }


                            string allKeywordsStr = keywordArray == null || keywordArray.Length == 0 ? "" : CommonUtil.StringArrayToString(keywordArray, ";");
                            Dictionary<string, object> keywordFileRow = new Dictionary<string, object>();
                            keywordFileRow.Add("publication", listRow["publication"]);
                            keywordFileRow.Add("host", listRow["host"]);
                            keywordFileRow.Add("title", listRow["title"]);
                            keywordFileRow.Add("authors", listRow["authors"]);
                            keywordFileRow.Add("year", year);
                            keywordFileRow.Add("pageUrl", pageUrl);
                            keywordFileRow.Add("keywords", allKeywordsStr);
                            keywordFileRow.Add("abstracts", listRow["abstracts"]);
                            keywordFileRow.Add("keywordsLength", allKeywordsStr.Length);
                            keywordFileRow.Add("txtUrl", keywordsTextFilePath);
                            allFileKeywordsWriter.AddRow(keywordFileRow);

                        }
                        catch (Exception ex)
                        {
                            string filePath = this.RunPage.GetFilePath(pageUrl, sourceDir);
                            this.RunPage.InvokeAppendLogText("错误，filePath = " + filePath + ", pageUrl = " + pageUrl, LogLevelType.Error, true);
                            //throw ex;

                            Dictionary<string, object> keywordFileRow = new Dictionary<string, object>();
                            keywordFileRow.Add("publication", listRow["publication"]);
                            keywordFileRow.Add("host", listRow["host"]);
                            keywordFileRow.Add("title", listRow["title"]);
                            keywordFileRow.Add("authors", listRow["authors"]);
                            keywordFileRow.Add("year", year);
                            keywordFileRow.Add("pageUrl", pageUrl);
                            keywordFileRow.Add("keywords", "");
                            keywordFileRow.Add("abstracts", listRow["abstracts"]);
                            keywordFileRow.Add("error", ex.Message);
                            keywordFileRow.Add("txtUrl", keywordsTextFilePath);
                            allFileKeywordsWriter.AddRow(keywordFileRow);
                        }
                    }
                }

                keywordsWriter.SaveToDisk();
                allFileKeywordsWriter.SaveToDisk();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        } 

        private string[] FindKeywords(string textFilePath)
        {
            string fileText = FileHelper.GetTextFromFile(textFilePath);
            string fileTextLower = fileText.ToLower();
            string keywordsName = "keywords:";
            int keywordStartIndex = fileTextLower.IndexOf(keywordsName);
            if(keywordStartIndex<0){
                keywordsName = "keyword:";
                keywordStartIndex = fileTextLower.IndexOf(keywordsName);
            }

            if (keywordStartIndex >= 0)
            {
                string partText = fileTextLower.Substring(keywordStartIndex + keywordsName.Length, 500);
                int abstractStartIndex = this.FindAbstract(partText);
                int jelStartIndex = this.FindJEL(partText);

                int allKeywordLength = 0;
                if (abstractStartIndex >= 0 || jelStartIndex >= 0)
                {
                    if (abstractStartIndex < 0)
                    {
                        allKeywordLength = jelStartIndex;
                    }
                    else if (jelStartIndex<0)
                    {
                        allKeywordLength = abstractStartIndex;
                    }
                    else
                    {
                        allKeywordLength = abstractStartIndex < jelStartIndex ? abstractStartIndex : jelStartIndex;
                    }
                }
                else
                {
                    int twoRNIndex = this.FindSectionTwoRN(partText);
                    if (twoRNIndex >= 0)
                    {
                        allKeywordLength = twoRNIndex;
                    }
                }

                if (allKeywordLength > 0)
                {
                    string keyWordsStr = fileText.Substring(keywordStartIndex + keywordsName.Length, allKeywordLength);
                    keyWordsStr = keyWordsStr.Trim();
                    if (keyWordsStr.StartsWith(":"))
                    {
                        keyWordsStr = keyWordsStr.Substring(1).Trim();
                    }
                    List<string> keywordList = new List<string>();
                    if (keyWordsStr.Contains(";") || keyWordsStr.Contains(","))
                    {
                        string[] keywordArray = keyWordsStr.Split(new string[] { ",", ";" }, StringSplitOptions.RemoveEmptyEntries);
                        for (int i = 0; i < keywordArray.Length; i++)
                        {
                            string keyword = keywordArray[i];
                            StringBuilder keywordBuilder = new StringBuilder();
                            for (int j = 0; j < keyword.Length; j++)
                            {
                                string k = keyword[j].ToString();
                                if (k == "\n")
                                {
                                    //忽略
                                }
                                else if (k == "\r")
                                {
                                    if (j > 0 && keyword[j - 1].ToString() != "-")
                                    {
                                        keywordBuilder.Append(" ");
                                    }
                                }
                                else if (k == "-")
                                {
                                    if (j < keyword.Length - 1 && keyword[j + 1].ToString() != "\r")
                                    {
                                        keywordBuilder.Append(k);
                                    }
                                }
                                else
                                {
                                    keywordBuilder.Append(k);
                                }
                                //keywordList.Add(keyword.Replace("\r\n", " ").Trim());
                            }
                            keywordList.Add(keywordBuilder.ToString().Trim());
                        }
                    }
                    else
                    {
                        string[] keywordArray = keyWordsStr.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
                        for (int i = 0; i < keywordArray.Length; i++)
                        {
                            string keyword = keywordArray[i];
                            keywordList.Add(keyword.Trim());
                        }
                    }
                    for (int i = 0; i < keywordList.Count; i++)
                    {
                        string keyword = keywordList[i];
                        keyword = keyword.TrimStart(new char[] { '”', '‘', '.' });
                        keyword = keyword.TrimEnd(new char[] { '”', '‘', '.' });
                        keywordList[i] = keyword;
                    }

                    return keywordList.ToArray();
                }
                else
                {
                    throw new Exception("Can not find Keywords End.");
                }
            }
            else
            {
                throw new Exception("Can not find KEYWORDS Begin.");
            }
        }

        private int FindAbstract(string partText)
        {
            int currentStartIndex = partText.IndexOf("a");
            while (currentStartIndex >= 0)
            {
                if (currentStartIndex >= 0)
                {
                    string trimNextPart = partText.Substring(currentStartIndex).Replace(" ", "");
                    if (trimNextPart.StartsWith("abstract"))
                    {
                        return currentStartIndex;
                    }
                    else
                    {
                        currentStartIndex = partText.IndexOf("a", currentStartIndex + 1);
                    }
                }
                else
                {
                    return -1;
                }
            }
            return -1;
        }

        private int FindJEL(string partText)
        {
            int currentStartIndex = partText.IndexOf("j");
            while (currentStartIndex >= 0)
            {
                if (currentStartIndex >= 0)
                {
                    string trimNextPart = partText.Substring(currentStartIndex).Replace(" ", "");
                    if (trimNextPart.StartsWith("jelno:")
                        || trimNextPart.StartsWith("jelclassification:")
                        || trimNextPart.StartsWith("jelclassifications:")
                        || trimNextPart.StartsWith("jelcodes:")
                        || trimNextPart.StartsWith("jel:"))
                    {
                        return currentStartIndex;
                    }
                    else
                    {
                        currentStartIndex = partText.IndexOf("j", currentStartIndex + 1);
                    }
                }
                else
                {
                    return -1;
                }
            }
            return -1;
        }

        private int FindSectionTwoRN(string partText)
        {
            int twoRNIndex = partText.IndexOf("\r\n\r\n");
            return twoRNIndex;
        }

        private int GetYear(string host)
        {
            string[] hostParts = host.Trim().Split(new string[] { " ", "\r\n", "<!-- -->" }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < hostParts.Length; i++)
            {
                string yearStr = hostParts[i];
                int year = 0;
                if (int.TryParse(yearStr, out year))
                {
                    if (year > 1900 && year < 2050)
                    {
                        return year;
                    }
                }
            }
            throw new Exception("无法找到年份, host = " + host);
        }

        private ExcelWriter GetDownloadPdfExcelWriter(string filePath)
        {
            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("publication", 0);
            resultColumnDic.Add("host", 1);
            resultColumnDic.Add("title", 2);
            resultColumnDic.Add("authors", 3);
            resultColumnDic.Add("abstracts", 4);
            resultColumnDic.Add("refs", 5);
            resultColumnDic.Add("pageUrl", 6);
            resultColumnDic.Add("txtUrl", 7);

            ExcelWriter resultEW = new ExcelWriter(filePath, "List", resultColumnDic, null);
            return resultEW;
        }

        private ExcelWriter GetKeywordsExcelWriter(string filePath)
        {
            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("publication", 0);
            resultColumnDic.Add("host", 1);
            resultColumnDic.Add("title", 2);
            resultColumnDic.Add("authors", 3);
            resultColumnDic.Add("year", 4);
            resultColumnDic.Add("keyword", 5);
            resultColumnDic.Add("pageUrl", 6);
            resultColumnDic.Add("txtUrl", 7);

            Dictionary<string, string> columnNameToFormats = new Dictionary<string, string>();
            columnNameToFormats.Add("year", "#0");

            ExcelWriter resultEW = new ExcelWriter(filePath, "List", resultColumnDic, columnNameToFormats);
            return resultEW;
        }

        private ExcelWriter GetAllFileKeywordsExcelWriter(string filePath)
        {
            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("publication", 0);
            resultColumnDic.Add("host", 1);
            resultColumnDic.Add("title", 2);
            resultColumnDic.Add("authors", 3);
            resultColumnDic.Add("year", 4);
            resultColumnDic.Add("keywords", 5);
            resultColumnDic.Add("abstracts", 6);
            resultColumnDic.Add("error", 7);
            resultColumnDic.Add("pageUrl", 8);
            resultColumnDic.Add("txtUrl", 9);
            resultColumnDic.Add("keywordsLength", 10);

            Dictionary<string, string> columnNameToFormats = new Dictionary<string, string>();
            columnNameToFormats.Add("year", "#0");
            columnNameToFormats.Add("keywordsLength", "#0");

            ExcelWriter resultEW = new ExcelWriter(filePath, "List", resultColumnDic, columnNameToFormats);
            return resultEW;
        } 
    }
}
