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
using NetDataAccess.Base.Reader;

namespace NetDataAccess.Extended.Jianzhu.QGJZSCJGGGFWPT_XM
{
    public class GetProjectAllListPage : ExternalRunWebPage
    {
        public override byte[] GetRequestData_BeforeSendRequest(string pageUrl, Dictionary<string, string> listRow, Encoding encoding)
        {
            string formData = listRow["formData"];
            if (formData != null && formData.Length > 0)
            {
                return encoding.GetBytes(formData);
            }
            else
            {
                return null;
            }
        }
        public override void WebRequestHtml_BeforeSendRequest(string pageUrl, Dictionary<string, string> listRow, Base.Web.NDAWebClient client)
        {
            client.Headers.Add("content-type", "application/x-www-form-urlencoded");
        }

        public override bool AfterAllGrab(IListSheet listSheet)
        {
            try
            {
                return this.GetAllPageUrls(listSheet);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private ExcelWriter GetExcelWriter(int fileIndex)
        {
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);

            string resultFilePath = Path.Combine(exportDir, "项目数据_项目信息_" + fileIndex.ToString() + ".xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }

        private bool GetAllPageUrls(IListSheet listSheet)
        {
            ExcelWriter ew = null;
            int projectIndex = 0;
            int fileIndex = 1;
            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;
            Dictionary<string, string> projectDic = new Dictionary<string, string>();
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailPageUrl = row[SysConfig.DetailPageUrlFieldName];
                string detailPageName = row[SysConfig.DetailPageNameFieldName];

                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    HtmlAgilityPack.HtmlDocument pageHtmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);

                    HtmlNodeCollection projectNodeList = pageHtmlDoc.DocumentNode.SelectNodes("//table[@class=\"table_box responsive personal\"]/tbody/tr");
                    if (projectNodeList != null)
                    {
                        for (int j = 0; j < projectNodeList.Count; j++)
                        {
                            HtmlNode projectNode = projectNodeList[j];
                            HtmlNodeCollection filedNodeList = projectNode.SelectNodes("./td");
                            if (filedNodeList[0].GetAttributeValue("data-header", "") == "序号")
                            {
                                HtmlNode linkNode = filedNodeList[2].SelectSingleNode("./a");
                                string linkValue = linkNode.GetAttributeValue("href", "");
                                int codeBeginIndex = linkValue.LastIndexOf("/") + 1;
                                string code = linkValue.Substring(codeBeginIndex).Trim();

                                if (!projectDic.ContainsKey(code))
                                {
                                    projectDic.Add(code, null);
                                    if (projectIndex % 500000 == 0)
                                    {
                                        if (ew != null)
                                        {
                                            ew.SaveToDisk();
                                        }
                                        ew = this.GetExcelWriter(fileIndex);
                                        fileIndex++;
                                    }
                                    Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                    f2vs.Add("detailPageUrl", "http://jzsc.mohurd.gov.cn/dataservice/query/project/projectDetail/" + code);
                                    f2vs.Add("detailPageName", code);
                                    ew.AddRow(f2vs);

                                    projectIndex++;
                                }
                            }
                        }
                    }
                }
            }

            ew.SaveToDisk();
            return true;
        }
    }
}