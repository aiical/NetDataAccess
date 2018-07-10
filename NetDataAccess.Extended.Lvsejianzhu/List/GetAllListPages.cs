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

namespace NetDataAccess.Extended.Lvsejianzhu
{
    /// <summary>
    /// GetAllListPages
    /// </summary>
    [System.Runtime.InteropServices.ComVisibleAttribute(true)]
    public class GetAllListPages : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            return this.GetShopInfos(listSheet);
        }

        private string GetShopUrl(string url)
        {
            int qIndex = url.IndexOf("?");
            if (qIndex >= 0)
            {
                return url.Substring(0, qIndex).Replace("/", "");
            }
            else
            {
                return url.Replace("/", "");
            }
        }

        private string GetShopId(string shopUrl)
        {
            int dIndex = shopUrl.IndexOf(".");
            return "https://" + shopUrl.Substring(0, dIndex);
        }

        private string GetShopType(string shopUrl)
        {
            int ldIndex = shopUrl.LastIndexOf(".");
            string tempStr = shopUrl.Substring(0, ldIndex);
            int ddIndex = tempStr.LastIndexOf(".");
            if (ddIndex >= 0)
            {
                return tempStr.Substring(ddIndex + 1).ToLower();
            }
            else
            {
                return tempStr.ToLower();
            }
        }

        private string GetShopProductListPageUrl(string shopId)
        {
            string url = "https://" + shopId + ".taobao.com/search.htm";
            return url;
        }

        /// <summary>
        /// 获取列表页里的店铺信息
        /// </summary>
        /// <param name="listSheet"></param>
        /// <returns></returns>
        private bool GetShopInfos(IListSheet listSheet)
        {
            bool succeed = true;
            string exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> columnDic = CommonUtil.InitStringIndexDic(new string[]{
                    "detailPageUrl",
                    "detailPageName", 
                    "cookie",
                    "grabStatus", 
                    "giveUpGrab",
                    "projectName",
                    "pageNum"});
            string shopFirstPageUrlFilePath = Path.Combine(exportDir, "项目详情.xlsx");
            ExcelWriter ew = new ExcelWriter(shopFirstPageUrlFilePath, "List", columnDic, null);

            Dictionary<string, string> projectUrlToNull = new Dictionary<string, string>();
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"]; 
                string pageNum = row["pageNum"]; 
                string localFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir);
                TextReader tr = null;

                try
                {
                    tr = new StreamReader(localFilePath, Encoding.GetEncoding(((Proj_Detail_SingleLine)this.RunPage.Project.DetailGrabInfoObject).Encoding));
                    string webPageHtml = tr.ReadToEnd();

                    HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                    htmlDoc.LoadHtml(webPageHtml);

                    HtmlNodeCollection listNodeList = htmlDoc.DocumentNode.SelectNodes("//div[@id=\"nyleft_down01\"]/table/tr/td/table/tr[1]/td");

                    this.GetProjectItem(listNodeList, pageNum, projectUrlToNull, ew); 
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
            ew.SaveToDisk();
            return succeed;
        }

        private void GetProjectItem(HtmlNodeCollection listNodeList, string pageNum, Dictionary<string, string> projectUrlToNull, ExcelWriter ew)
        {
            for (int j = 0; j < listNodeList.Count; j++)
            {
                HtmlNode listNode = listNodeList[j];

                string projectName = "";
                string projectUrl = ""; 
                
                HtmlNode projectNameNode = listNode.SelectSingleNode("./a[1]");
                projectName = projectNameNode.InnerText.Trim();

                HtmlNode projectUrlNode = listNode.SelectSingleNode("./a[2]"); 
                projectUrl = "http://www.gbmap.org" + projectUrlNode.GetAttributeValue("href", "");
                 
                Dictionary<string, object> projectInfo = new Dictionary<string, object>();
                if (!projectUrlToNull.ContainsKey(projectUrl))
                {
                    projectUrlToNull.Add(projectUrl, null);
                    projectInfo.Add("detailPageUrl", projectUrl);
                    projectInfo.Add("detailPageName", projectUrl);
                    projectInfo.Add("projectName", projectName);
                    projectInfo.Add("pageNum", pageNum);
                    ew.AddRow(projectInfo);
                }
            }
        } 
        public override void CheckRequestCompleteFile(string webPageText, Dictionary<string, string> listRow)
        {
            if (!webPageText.Contains("下一页"))
            {
                throw new Exception("Uncompleted webpage request.");
            } 
        }
    }
}