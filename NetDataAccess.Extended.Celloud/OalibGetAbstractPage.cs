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
using System.Web;

namespace NetDataAccess.Extended.Celloud
{
    public class OalibGetAbstractPage : CustomProgramBase
    {
        public bool Run(string parameters, IListSheet listSheet)
        {
            return GetAllListPageUrl(listSheet)
                && GetAllUrlHost(listSheet);
        }
        private bool GetAllListPageUrl(IListSheet listSheet)
        {
            string exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("species", 5); 
            resultColumnDic.Add("year", 7);
            resultColumnDic.Add("abstractUrl", 8);
            resultColumnDic.Add("code", 9);
            resultColumnDic.Add("paperName", 10);
            string resultFilePath = Path.Combine(exportDir, "oalib获取论文Html页.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;
            string detailPageNameColumnName = SysConfig.DetailPageNameFieldName;

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string url = row[detailPageUrlColumnName];
                    string code = row[detailPageNameColumnName];
                    string cookie = row["cookie"];
                    string species = row["species"].Trim(); 
                    string year = row["year"].Trim();
                    string paperName = row["paperName"];
                    string localFilePath = this.RunPage.GetFilePath(url, pageSourceDir);
                    TextReader tr = null;

                    try
                    {
                        tr = new StreamReader(localFilePath);
                        string webPageHtml = tr.ReadToEnd();

                        HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                        htmlDoc.LoadHtml(webPageHtml);

                        HtmlNodeCollection allPaperNodes = htmlDoc.DocumentNode.SelectNodes("//p[@class=\"resetHref\"]/a");
                        if (allPaperNodes != null)
                        {
                            HtmlNode paperLinkNode = null;
                            foreach (HtmlNode paperNode in allPaperNodes)
                            {
                                string paperLinkText = paperNode.InnerText.Trim();
                                if (paperLinkText == "Full-Text")
                                {
                                    paperLinkNode = paperNode;
                                    break;
                                }
                            }
                            if (paperLinkNode == null)
                            {  
                                this.RunPage.InvokeAppendLogText("警告: 未成功定位到找到论文页地址链接.url = " + url, LogLevelType.Warring, true);
                                //throw new Exception("未成功定位到找到论文页地址链接.url = " + url);
                            }
                            else
                            {
                                string paperUrl = CommonUtil.HtmlDecode(paperLinkNode.Attributes["href"].Value);
                                Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                f2vs.Add("detailPageUrl", paperUrl);
                                f2vs.Add("detailPageName", code);
                                f2vs.Add("cookie", cookie);
                                f2vs.Add("paperName", paperName);
                                f2vs.Add("species", species); 
                                f2vs.Add("year", year);
                                f2vs.Add("code", code);
                                f2vs.Add("abstractUrl", url);
                                resultEW.AddRow(f2vs);
                            }
                        }
                        else
                        {

                            throw new Exception("无法找到论文页地址链接.url = " + url);
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
        private bool GetAllUrlHost(IListSheet listSheet)
        {
            string exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            
            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("hostName", 0);
            resultColumnDic.Add("fileType", 1);
            resultColumnDic.Add("canProcess", 2);
            resultColumnDic.Add("exampleUrl", 3);
            resultColumnDic.Add("paperContentXPath", 4);
            string resultFilePath = Path.Combine(exportDir, "oalib论文格式分类.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;
            string detailPageNameColumnName = SysConfig.DetailPageNameFieldName;
            List<string> allHosts= new  List<string>();
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string url = row[detailPageUrlColumnName];
                    Uri uri = new Uri(url);
                    string hostName = uri.Host;
                    if (!allHosts.Contains(hostName))
                    {
                        allHosts.Add(hostName);
                        Dictionary<string, string> f2vs = new Dictionary<string, string>();
                        f2vs.Add("hostName", hostName);
                        f2vs.Add("exampleUrl", url);
                        resultEW.AddRow(f2vs);
                    }
                }
            }
            resultEW.SaveToDisk();

            return true;
        }
    }
}