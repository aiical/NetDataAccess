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
    public class AtlassianPageInfo : CustomProgramBase
    { 
        public bool Run(string parameters, IListSheet listSheet )
        {
            try
            {
                return this.GenerateAtlassianPageInfo(listSheet); 
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        private bool GenerateAtlassianPageInfo(IListSheet listSheet)
        {
            bool succeed = true;
            string exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> pageInfoColumnDic = new Dictionary<string, int>();
            pageInfoColumnDic.Add("code", 0);
            pageInfoColumnDic.Add("url", 1);
            pageInfoColumnDic.Add("title", 2);
            pageInfoColumnDic.Add("assigneeCode", 3);
            pageInfoColumnDic.Add("assigneeName", 4);
            pageInfoColumnDic.Add("reporterCode", 5);
            pageInfoColumnDic.Add("reporterName", 6);
            pageInfoColumnDic.Add("testedUserCode", 7);
            pageInfoColumnDic.Add("testedUserName", 8);
            pageInfoColumnDic.Add("type", 9);
            pageInfoColumnDic.Add("priority", 10);
            pageInfoColumnDic.Add("status", 11);
            pageInfoColumnDic.Add("resolution", 12);
            string readDetailDir = this.RunPage.GetReadFileDir();
            string pageInfoFilePath = Path.Combine(exportDir, this.RunPage.Project.Name + "_info.xlsx");
            
            string outputTitleTextDir = Path.Combine(exportDir, "titleText");

            ExcelWriter pageInfoEW = new ExcelWriter(pageInfoFilePath, "List", pageInfoColumnDic); 
            for (int i = 0; i < listSheet.RowCount ; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string code = row["title"];
                    string url = row["detailPageUrl"];
                    string title = "";
                    string assigneeCode = "";
                    string assigneeName = "";
                    string reporterCode="";
                    string reporterName="";
                    string testedUserCode="";
                    string testedUserName="";
                    string type="";
                    string priority="";
                    string status="";
                    string resolution = "";

                    string localFilePath = this.RunPage.GetFilePath(url, pageSourceDir);
                    TextReader tr = null;

                    try
                    {
                        tr = new StreamReader(localFilePath);
                        string webPageHtml = tr.ReadToEnd();

                        HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                        htmlDoc.LoadHtml(webPageHtml);

                        {
                            HtmlNode titleNode = htmlDoc.DocumentNode.SelectSingleNode("//*[@id=\"summary-val\"]");
                            title = HttpUtility.HtmlDecode(titleNode.InnerText.Trim());
                        }
                        {
                            HtmlNode assigneeCodeNode = htmlDoc.DocumentNode.SelectSingleNode("//*[@id=\"peopledetails\"]/li[1]/dl[1]/dd[1]/span[1]/span[1]/span[1]/span[1]/img[1]");
                            if (assigneeCodeNode == null)
                            {
                                assigneeCodeNode = htmlDoc.DocumentNode.SelectSingleNode("//*[@id=\"peopledetails\"]/li[1]/dl[1]/dd[1]/span[1]/span[1]/span[1]/img[1]");
                            }
                            HtmlAttribute assigneeCodeAttri = assigneeCodeNode.Attributes["alt"];
                            assigneeCode = assigneeCodeAttri == null ? "" : assigneeCodeAttri.Value.Trim();
                        }
                        {
                            HtmlNode assigneeNameNode = htmlDoc.DocumentNode.SelectSingleNode("//*[@id=\"peopledetails\"]/li[1]/dl[1]/dd[1]/span[1]/span[1]");
                            assigneeName = HttpUtility.HtmlDecode(assigneeNameNode.InnerText.Trim());
                        }
                        {
                            HtmlNode reporterCodeNode = htmlDoc.DocumentNode.SelectSingleNode("//*[@id=\"peopledetails\"]/li[1]/dl[2]/dd[1]/span[1]/span[1]/span[1]/span[1]/img[1]");
                            if (reporterCodeNode == null)
                            {
                                reporterCodeNode = htmlDoc.DocumentNode.SelectSingleNode("//*[@id=\"peopledetails\"]/li[1]/dl[2]/dd[1]/span[1]/span[1]/span[1]/img[1]");
                            }
                            HtmlAttribute reporterCodeAttri = reporterCodeNode.Attributes["alt"];
                            reporterCode = reporterCodeAttri == null ? "" : reporterCodeAttri.Value.Trim();
                        }
                        {
                            HtmlNode reporterNameNode = htmlDoc.DocumentNode.SelectSingleNode("//*[@id=\"peopledetails\"]/li[1]/dl[2]/dd[1]/span[1]/span[1]");
                            reporterName = HttpUtility.HtmlDecode(reporterNameNode.InnerText.Trim());
                        }
                        {
                            HtmlNode testedUserNode = htmlDoc.DocumentNode.SelectSingleNode("//*[@id=\"peopledetails\"]/li[1]/dl[3]/dd[1]/span[1]/div[1]/span[1]/span[1]");
                            if (testedUserNode == null)
                            {
                                testedUserNode = htmlDoc.DocumentNode.SelectSingleNode("//*[@id=\"peopledetails\"]/li[1]/dl[3]/dd[1]/span[1]/div[1]/span[1]");
                            }
                            testedUserCode = testedUserNode == null ? "" : testedUserNode.Attributes["rel"].Value;
                            testedUserName = HttpUtility.HtmlDecode(testedUserNode == null ? "" : testedUserNode.InnerText.Trim());
                        }
                        {
                            type = HttpUtility.HtmlDecode(htmlDoc.DocumentNode.SelectSingleNode("//*[@id=\"type-val\"]").InnerText.Trim());
                            priority = HttpUtility.HtmlDecode(htmlDoc.DocumentNode.SelectSingleNode("//*[@id=\"priority-val\"]").InnerText.Trim());
                            status = HttpUtility.HtmlDecode(htmlDoc.DocumentNode.SelectSingleNode("//*[@id=\"status-val\"]").InnerText.Trim());
                            resolution =HttpUtility.HtmlDecode( htmlDoc.DocumentNode.SelectSingleNode("//*[@id=\"resolution-val\"]").InnerText.Trim());
                        }

                        Dictionary<string, string> f2vs = new Dictionary<string, string>(); 
                        f2vs.Add("code", code); 
                        f2vs.Add("url", url); 
                        f2vs.Add("title", title); 
                        f2vs.Add("assigneeCode", assigneeCode); 
                        f2vs.Add("assigneeName", assigneeName); 
                        f2vs.Add("reporterCode", reporterCode); 
                        f2vs.Add("reporterName", reporterName); 
                        f2vs.Add("testedUserCode", testedUserCode); 
                        f2vs.Add("testedUserName", testedUserName); 
                        f2vs.Add("type", type); 
                        f2vs.Add("priority", priority); 
                        f2vs.Add("status", status); 
                        f2vs.Add("resolution", resolution);

                        //保存标题文本
                        string titleTextFilePath = Path.Combine(outputTitleTextDir, code + ".txt");
                        CommonUtil.CreateFileDirectory(titleTextFilePath);

                        TextWriter tw = null;
                        try
                        {
                            tw = new StreamWriter(titleTextFilePath, false, new UTF8Encoding(false));
                            tw.Write(title);
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
                        pageInfoEW.AddRow(f2vs);

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
            pageInfoEW.SaveToDisk();
            return succeed;
        } 
    }
}