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
    public class IDempiereListPage : CustomProgramBase
    {
        public bool Run(string parameters, IListSheet listSheet)
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

            Dictionary<string, int> subjectColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                "detailPageUrl", 
                "detailPageName", 
                "cookie", 
                "grabStatus",
                "giveUpGrab",
                "subjectIndex",
                "id",
                "title",
                "googleUrl",
                "creator",
                "createDate",
                "messageCount", 
                "authorCount" });
            string subjectFileExcelPath = Path.Combine(exportDir, this.RunPage.Project.Name + "_List.xlsx");
            ExcelWriter subjectEW = new ExcelWriter(subjectFileExcelPath, "List", subjectColumnDic);

            Dictionary<string, int> subjectImportColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                "index",
                "id",
                "creator",
                "createDate", 
                "commentPublisher",
                "commentPublished",
                "googleUrl",
                "title", 
                "url"});
            string subjectFileXmlPath = Path.Combine(exportDir, this.RunPage.Project.Name + "_List.xml");
            XmlWriter subjectXW = new XmlWriter(subjectFileXmlPath, subjectImportColumnDic);

            string[] commentUsers = new string[] { "sunhua", "shizhengzhong", "liyuzhu" };

            int sujectIndex = 1;
            for (int i = listSheet.RowCount - 1; i >= 0; i--)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string url = row[SysConfig.DetailPageUrlFieldName];
                    string localFilePath = this.RunPage.GetFilePath(url, pageSourceDir);
                    TextReader tr = null;

                    try
                    {
                        tr = new StreamReader(localFilePath);
                        string webPageHtml = tr.ReadToEnd();

                        HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                        htmlDoc.LoadHtml(webPageHtml);
                        HtmlNodeCollection subjectNodes = htmlDoc.DocumentNode.SelectNodes("//body[1]/table[1]/tr");

                        if (subjectNodes != null)
                        {
                            for (int j = subjectNodes.Count - 1; j >= 1; j--)
                            {
                                HtmlNode subjectNode = subjectNodes[j];
                                HtmlNode titleNode = subjectNode.SelectSingleNode("./td[@class=\"subject\"]/a");
                                string title = CommonUtil.HtmlDecode(CommonUtil.ReplaceAsciiByString(titleNode.InnerText.Trim()));
                                string detailPageUrl = titleNode.Attributes["href"].Value.Trim();
                                string detailPageName = detailPageUrl.Substring(detailPageUrl.LastIndexOf("/") + 1);
                                HtmlNode authorNode = subjectNode.SelectSingleNode("./td[@class=\"author\"]");
                                string author = authorNode == null ? "" : CommonUtil.HtmlDecode(CommonUtil.ReplaceAsciiByString(authorNode.InnerText.Trim()));
                                HtmlNode lastPostDateNode = subjectNode.SelectSingleNode("./td[@class=\"lastPostDate\"]");
                                string lastPostDate = lastPostDateNode == null ? "" : lastPostDateNode.InnerText.Trim();

                                Dictionary<string, string> commentF2vs = new Dictionary<string, string>();
                                commentF2vs.Add("subjectIndex", sujectIndex.ToString());
                                commentF2vs.Add("detailPageUrl", detailPageUrl);
                                commentF2vs.Add("detailPageName", detailPageName);
                                commentF2vs.Add("id", detailPageName);
                                commentF2vs.Add("title", title); 
                                commentF2vs.Add("googleUrl", detailPageUrl);
                                commentF2vs.Add("creator", author);
                                commentF2vs.Add("createDate", lastPostDate);
                                subjectEW.AddRow(commentF2vs);

                                for (int u = 0; u < commentUsers.Length; u++)
                                {
                                    string user = commentUsers[u];
                                    Dictionary<string, string> commentXF2vs = new Dictionary<string, string>();
                                    commentXF2vs.Add("index", sujectIndex.ToString());
                                    commentXF2vs.Add("id", user + "_" + detailPageName);
                                    commentXF2vs.Add("title", title);
                                    commentXF2vs.Add("url", detailPageName + ".html");
                                    commentXF2vs.Add("googleUrl", detailPageUrl);
                                    commentXF2vs.Add("creator", author);
                                    commentXF2vs.Add("createDate", lastPostDate);
                                    commentXF2vs.Add("commentPublisher", user);
                                    commentXF2vs.Add("commentPublished", "No");
                                    subjectXW.AddRow(commentXF2vs);
                                }

                                sujectIndex++;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        if (tr != null)
                        {
                            tr.Close();
                            tr.Dispose();

                        }
                        throw ex;
                    }
                }
            }
            subjectXW.SaveToDisk();
            subjectEW.SaveToDisk();
            return succeed;
        }
    }
}