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
    public class IDPageList : CustomProgramBase
    { 
        public bool Run(string parameters, IListSheet listSheet )
        {
            try
            {
                return this.GenerateNewPage(listSheet); 
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        private bool GenerateNewPage(IListSheet listSheet)
        {
            bool succeed = true;
            string exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            string outputTitleTextDir = Path.Combine(exportDir, "titleText");

            Dictionary<string, int> subjectColumnDic =  CommonUtil.InitStringIndexDic(new string[]{
                "index",
                "id",
                "title",
                "url", 
                "googleUrl",
                "creator",
                "createDate",
                "messageCount", 
                "authorCount" ,
                "commentPublisher",
                "commentPublished"});
            string subjectFileExcelPath = Path.Combine(exportDir, this.RunPage.Project.Name + "_List.xlsx");
            ExcelWriter subjectEW = new ExcelWriter(subjectFileExcelPath, "List", subjectColumnDic);

            string subjectFileXmlPath = Path.Combine(exportDir, this.RunPage.Project.Name + "_List.xml");
            XmlWriter subjectXW = new XmlWriter(subjectFileXmlPath, subjectColumnDic);


            Dictionary<string, int> allCommentsColumnDic = new Dictionary<string, int>();
            allCommentsColumnDic.Add("subjectIndex", 0);
            allCommentsColumnDic.Add("googleUrl", 1);
            allCommentsColumnDic.Add("creator", 2); 
            allCommentsColumnDic.Add("author", 3);
            allCommentsColumnDic.Add("lastPostDate", 4);
            string allCommentsFilePath = Path.Combine(exportDir, this.RunPage.Project.Name + "_AllComments.xlsx");
            ExcelWriter allCommentsListEW = new ExcelWriter(allCommentsFilePath, "List", allCommentsColumnDic);
              
            for (int i = 0; i < listSheet.RowCount ; i++)
            {
                Dictionary<string,string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string index = row[SysConfig.ListPageIndexFieldName].PadLeft(4, '0');
                    string googleUrl = row[SysConfig.DetailPageUrlFieldName];
                    string title = row["title"];
                    string creator = row["creator"];
                    string createDate = row["createDate"];
                    string localFilePath = this.RunPage.GetFilePath(googleUrl, pageSourceDir);
                    TextReader tr = null;

                    try
                    {
                        tr = new StreamReader(localFilePath);
                        string webPageHtml = tr.ReadToEnd();

                        HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                        htmlDoc.LoadHtml(webPageHtml);

                        string messagesNumStr = htmlDoc.DocumentNode.SelectSingleNode("//body/i").InnerText;
                        int ofIndex = messagesNumStr.IndexOf("of");
                        int messagesIndex = messagesNumStr.IndexOf("messages");
                        string messageCount = messagesNumStr.Substring(ofIndex+2, messagesIndex-ofIndex-2).Trim();

                        List<string> authorNames = new List<string>();

                        HtmlNodeCollection messageNodes = htmlDoc.DocumentNode.SelectNodes("//body/table/tr");

                        if (messageNodes != null)
                        {
                            for (int j = 0; j < messageNodes.Count;j++ )
                            {
                                HtmlNode messageNode = messageNodes[j];
                                HtmlNode authorNode = messageNode.SelectSingleNode("./td[2]");
                                HtmlNode lastPostDateNode = messageNode.SelectSingleNode("./td[3]");
                                if (authorNode != null)
                                {
                                    string author = HttpUtility.HtmlDecode(authorNode.InnerText).Trim();
                                    if (j == 0)
                                    {
                                        creator = author;
                                    }
                                    if (!authorNames.Contains(author))
                                    {
                                        authorNames.Add(author);
                                    }
                                    string lastPostDate = lastPostDateNode == null ? "" : lastPostDateNode.InnerText;

                                    Dictionary<string, string> commentF2vs = new Dictionary<string, string>();
                                    commentF2vs.Add("subjectIndex",  index);
                                    commentF2vs.Add("googleUrl", googleUrl);
                                    commentF2vs.Add("creator", creator);
                                    commentF2vs.Add("author", author);
                                    commentF2vs.Add("lastPostDate", lastPostDate);
                                    allCommentsListEW.AddRow(commentF2vs);
                                }
                            }
                        }

                        //修改html内容，增加线框
                        HtmlNode tableNode = htmlDoc.DocumentNode.SelectSingleNode("//body/table");
                        tableNode.Attributes["border"].Value = "1";
                        int localHtmlUrlStartIndex = googleUrl.IndexOf("/idempiere/") + "/idempiere/".Length;
                        string htmlLocalName = CommonUtil.ProcessFileName(googleUrl.Substring(localHtmlUrlStartIndex), "_") + ".html";
                        string htmlLocalUrl = Path.Combine(Path.GetDirectoryName(pageSourceDir), "export\\html\\" + htmlLocalName);
                        CommonUtil.CreateFileDirectory(htmlLocalUrl);
                        htmlDoc.Save(htmlLocalUrl);

                        Dictionary<string, string> f2vs = new Dictionary<string, string>(); 
                        f2vs.Add("index", index);
                        f2vs.Add("title", title);
                        f2vs.Add("googleUrl", googleUrl);
                        f2vs.Add("url", htmlLocalName); 
                        f2vs.Add("creator", creator);
                        f2vs.Add("createDate", createDate);
                        f2vs.Add("messageCount", messageNodes.Count.ToString());
                        f2vs.Add("authorCount", authorNames.Count.ToString());
                        f2vs.Add("commentPublisher", "");
                        f2vs.Add("commentPublished", "");
                        f2vs.Add("id", "");

                        IRow newPageListRow = subjectEW.AddRow(f2vs);

                        f2vs["commentPublished"] = "No";

                        f2vs["commentPublisher"] = "sunhua";
                        f2vs["id"] = "sunhua" + index;
                        subjectXW.AddRow(f2vs);

                        f2vs["commentPublisher"] = "shizhengzhong";
                        f2vs["id"] = "shizhengzhong" + index;
                        subjectXW.AddRow(f2vs);

                        f2vs["commentPublisher"] = "liyuzhu";
                        f2vs["id"] = "liyuzhu" + index;
                        subjectXW.AddRow(f2vs);

                        ICell localUrlCell = subjectEW.GetCell(newPageListRow, "url", true);
                        IHyperlink hyperlink = new XSSFHyperlink( HyperlinkType.File);
                        hyperlink.Address = "html/" + htmlLocalName;
                        localUrlCell.Hyperlink = hyperlink;

                        //保存message文本
                        string msg = Encoding.UTF8.GetString(Encoding.Convert(Encoding.ASCII, Encoding.UTF8, Encoding.ASCII.GetBytes(title))).Trim();

                        if (msg.Length > 0)
                        { 
                            string titleTextFilePath = Path.Combine(outputTitleTextDir, index + ".txt");
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
                    catch (Exception ex)
                    {
                        if (tr != null)
                        {
                            tr.Dispose();
                            tr = null;
                        }
                        this.RunPage.InvokeAppendLogText("读取出错. " + ex.Message + " LocalPath = " + localFilePath, LogLevelType.Error, true);
                    }
                }
            } 
            subjectXW.SaveToDisk();
            subjectEW.SaveToDisk();
            allCommentsListEW.SaveToDisk();
            return succeed;
        } 
    }
}