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
using NetDataAccess.Extended.Taobao.Common;

namespace NetDataAccess.Extended.Id.List
{
    /// <summary>
    /// GetAllListPages
    /// </summary>
    [System.Runtime.InteropServices.ComVisibleAttribute(true)]
    public class GetAllListPages : ExternalRunWebPage
    {
        public override void CheckRequestCompleteFile(string webPageText, Dictionary<string, string> listRow)
        {
            base.CheckRequestCompleteFile(webPageText, listRow);
            TextReader tr = null;

            try
            {
                HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                htmlDoc.LoadHtml(webPageText);
                HtmlNodeCollection listNodeListA = htmlDoc.DocumentNode.SelectNodes("//ol[@class=\"issue-list\"]/li/a");
                if (listNodeListA == null || listNodeListA.Count == 0)
                {
                    throw new Exception("文档未加载完成");
                }
            }
            catch (Exception ex)
            {
                if (tr != null)
                {
                    tr.Dispose();
                    tr = null;
                }
                throw new Exception("文档未加载完成");
            }
        }

        public override bool AfterAllGrab(IListSheet listSheet)
        {
            return this.GetShopInfos(listSheet);
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
                    "code",
                    "name"});
            string shopFirstPageUrlFilePath = Path.Combine(exportDir, "Id详情页.xlsx");
            ExcelWriter ew = new ExcelWriter(shopFirstPageUrlFilePath, "List", columnDic, null);

            Dictionary<string, string> codeToNames = new Dictionary<string, string>();
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"]; 
                string localFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir);
                TextReader tr = null;

                try
                {
                    tr = new StreamReader(localFilePath, Encoding.GetEncoding(((Proj_Detail_SingleLine)this.RunPage.Project.DetailGrabInfoObject).Encoding));
                    string webPageHtml = tr.ReadToEnd();

                    HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                    htmlDoc.LoadHtml(webPageHtml);

                    HtmlNodeCollection listNodeListA = htmlDoc.DocumentNode.SelectNodes("//ol[@class=\"issue-list\"]/li/a");
                    if (listNodeListA.Count > 0 )
                    {
                        foreach (HtmlNode aNode in listNodeListA)
                        {
                            this.GetShopItem(aNode, codeToNames, ew);
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
            ew.SaveToDisk();
            return succeed;
        }

        private void GetShopItem(HtmlNode aNode, Dictionary<string, string> codeToNames, ExcelWriter ew)
        {
            string code = CommonUtil.HtmlDecode(aNode.SelectSingleNode("./span[@class=\"issue-link-key\"]").InnerText).Trim();
            string name = CommonUtil.HtmlDecode(aNode.SelectSingleNode("./span[@class=\"issue-link-summary\"]").InnerText).Trim();
            string url = "https://idempiere.atlassian.net" + aNode.GetAttributeValue("href", "");

            Dictionary<string, object> itemInfo = new Dictionary<string, object>();
            if (!codeToNames.ContainsKey(code))
            {
                codeToNames.Add(code, name);
                itemInfo.Add("detailPageUrl", url);
                itemInfo.Add("detailPageName", code);
                itemInfo.Add("code", code);
                itemInfo.Add("name", name);
                ew.AddRow(itemInfo);
            }
        } 
    }
}