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
using NPOI.XSSF.UserModel;
using NetDataAccess.Base.DB;
using System.Web; 

namespace NetDataAccess.Extended.IDempiere
{
    /// <summary>
    /// 处理页面
    /// </summary>
    public class BitbucketListPageInfo : CustomProgramBase
    { 
        public bool Run(string parameters, IListSheet listSheet )
        {
            try
            {
                return this.GenerateBitbucketListPageInfo(listSheet); 
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        private bool GenerateBitbucketListPageInfo(IListSheet listSheet)
        {
            bool succeed = true;
            string exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            string outputTitleTextDir = Path.Combine(exportDir, "titleText");

            Dictionary<string, int> pageInfoColumnDic = new Dictionary<string, int>();
            pageInfoColumnDic.Add("pageCode", 0);
            pageInfoColumnDic.Add("pageUrl", 1);
            pageInfoColumnDic.Add("authorCode", 2);
            pageInfoColumnDic.Add("authorName", 3);
            pageInfoColumnDic.Add("merge", 4);
            pageInfoColumnDic.Add("followCode", 5);
            pageInfoColumnDic.Add("message", 6);
            pageInfoColumnDic.Add("date", 7); 
            string readDetailDir = this.RunPage.GetReadFileDir();
            string pageListFilePath = Path.Combine(exportDir, this.RunPage.Project.Name + "_info.xlsx");
            ExcelWriter pageListEW = new ExcelWriter(pageListFilePath, "List", pageInfoColumnDic); 
            for (int i = 0; i < listSheet.RowCount ; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string listUrl = row["detailPageUrl"];
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {

                    string localFilePath = this.RunPage.GetFilePath(listUrl, pageSourceDir);
                    TextReader tr = null;

                    try
                    {
                        tr = new StreamReader(localFilePath);
                        string webPageHtml = tr.ReadToEnd();

                        HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                        htmlDoc.LoadHtml(webPageHtml);
                        HtmlNodeCollection trNodes = htmlDoc.DocumentNode.SelectNodes("//*[@id=\"commit-list-container\"]/table[1]/tbody[1]/tr");
                        foreach (HtmlNode trNode in trNodes)
                        {
                            HtmlNode pageNode = trNode.SelectSingleNode("./td[@class=\"hash\"]/div[1]/a[@class=\"hash execute\"]");
                            string pageCode = pageNode.InnerText.Trim();
                            string pageUrl = pageNode.Attributes["href"].Value;

                            HtmlNode authorNode = trNode.SelectSingleNode("./td[@class=\"user\"]/div[@class=\"author\"]/span[@class=\"author\"]/span[1]/a[1]");
                            string authorCodeStr = authorNode == null ? "" : authorNode.Attributes["href"].Value;
                            string authorCode = authorNode == null ? "" : authorCodeStr.Substring(1, authorCodeStr.Length - 2);
                            string authorName = authorNode == null ? "" : authorNode.Attributes["title"].Value;

                            HtmlNode mergeNode = trNode.SelectSingleNode("./td[@class=\"hash\"]/div[1]/span[@class=\"aui-lozenge\"]");
                            string merge = mergeNode == null ? "" : mergeNode.InnerText.Trim();

                            HtmlNodeCollection followNodes = trNode.SelectNodes("./td[@class=\"text flex-content--column\"]/div[@class=\"flex-content\"]/div[@class=\"flex-content--primary\"]/span[@class=\"subject\"]/a");
                            StringBuilder followCodeSB = new StringBuilder();
                            if (followNodes != null)
                            {
                                foreach (HtmlNode followNode in followNodes)
                                {
                                    string followCode = followNode.InnerText.Trim();
                                    followCodeSB.Append(followCode + ";");
                                }
                            }

                            HtmlNode messageNode = trNode.SelectSingleNode("./td[@class=\"text flex-content--column\"]/div[@class=\"flex-content\"]/div[@class=\"flex-content--primary\"]/span[@class=\"subject\"]");
                            string message = messageNode == null ? "" : HttpUtility.HtmlDecode(messageNode.InnerText).Trim();

                            HtmlNode dateNode = trNode.SelectSingleNode("./td[@class=\"date\"]/div[1]/time[1]");
                            string date = dateNode.Attributes["datetime"].Value;

                            Dictionary<string, string> f2vs = new Dictionary<string, string>();
                            f2vs.Add("pageCode", pageCode);
                            f2vs.Add("pageUrl", pageUrl);
                            f2vs.Add("authorCode", authorCode);
                            f2vs.Add("authorName", authorName);
                            f2vs.Add("merge", merge);
                            f2vs.Add("followCode", followCodeSB.ToString());
                            f2vs.Add("message", message);
                            f2vs.Add("date", date);
                            pageListEW.AddRow(f2vs);

                            //保存message文本
                            string msg = Encoding.UTF8.GetString(Encoding.Convert(Encoding.ASCII, Encoding.UTF8, Encoding.ASCII.GetBytes(message))).Trim();

                            if (msg.Length > 0)
                            { 
                                string titleTextFilePath = Path.Combine(outputTitleTextDir, pageCode + ".txt");
                                CommonUtil.CreateFileDirectory(titleTextFilePath);

                                TextWriter tw = null;
                                try
                                {
                                    tw = new StreamWriter(titleTextFilePath, false, new UTF8Encoding(false));
                                    tw.Write(msg);
                                    tw.Flush();
                                }
                                catch (Exception ee)
                                {
                                    throw ee;
                                }
                                finally
                                {
                                    if (tw != null)
                                    {
                                        tw.Close();
                                        tw.Dispose();
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        if (tr != null)
                        {
                            tr.Dispose();
                            tr = null;
                        }
                        this.RunPage.InvokeAppendLogText("GenerateGoods读取出错.  " + ex.Message + " LocalPath = " + localFilePath, LogLevelType.Error, true);
                    }
                }
            }
            pageListEW.SaveToDisk();
            return succeed;
        } 
    }
}