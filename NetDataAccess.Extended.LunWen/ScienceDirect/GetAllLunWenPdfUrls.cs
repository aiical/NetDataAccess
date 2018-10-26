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
            /*
            try
            {
                string sourceDir = this.RunPage.GetDetailSourceFileDir();
                string exportDir = this.RunPage.GetExportDir();
                string pdfUrlFilePath = Path.Combine(exportDir, "论文_ScienceDirect_论文PDF页.xlsx");
                ExcelWriter pdfUrlWriter = this.GetDownloadPdfExcelWriter(pdfUrlFilePath); 

                for (int i = 0; i < listSheet.RowCount; i++)
                {
                    Dictionary<string, string> listRow = listSheet.GetRow(i);
                    string pageUrl = listRow[SysConfig.DetailPageUrlFieldName];

                    try
                    {

                        HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);

                        HtmlNode linkNode = htmlDoc.DocumentNode.SelectSingleNode("//div[@id=\"redirect-message\"]/p/a");
                        string pdfUrl = linkNode.GetAttributeValue("href", "");

                        Dictionary<string, string> pdfUrlRow = new Dictionary<string, string>();
                        pdfUrlRow.Add("detailPageUrl", pdfUrl);
                        pdfUrlRow.Add("detailPageName", pdfUrl);
                        pdfUrlRow.Add("publication", listRow["publication"]);
                        pdfUrlRow.Add("host", listRow["host"]);
                        pdfUrlRow.Add("title", listRow["title"]);
                        pdfUrlRow.Add("authors", listRow["authors"]);
                        pdfUrlRow.Add("abstracts", listRow["abstracts"]);
                        pdfUrlRow.Add("refs", listRow["refs"]);
                        pdfUrlRow.Add("url", listRow["url"]);
                        pdfUrlWriter.AddRow(pdfUrlRow);
                    }
                    catch (Exception ex)
                    {
                        string filePath = this.RunPage.GetFilePath(pageUrl, sourceDir);
                        this.RunPage.InvokeAppendLogText("错误，filePath = " + filePath + ", pageUrl = " + pageUrl, LogLevelType.Error, true);
                        throw ex;
                    }
                }

                pdfUrlWriter.SaveToDisk(); 
                return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
             */
            return true;
        }

        private ExcelWriter GetDownloadPdfExcelWriter(string filePath)
        {
            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("publication", 5);
            resultColumnDic.Add("host", 6);
            resultColumnDic.Add("title", 7);
            resultColumnDic.Add("authors", 8);
            resultColumnDic.Add("abstracts", 9);
            resultColumnDic.Add("refs", 10);
            resultColumnDic.Add("url", 11);

            ExcelWriter resultEW = new ExcelWriter(filePath, "List", resultColumnDic, null);
            return resultEW;
        } 
    }
}
