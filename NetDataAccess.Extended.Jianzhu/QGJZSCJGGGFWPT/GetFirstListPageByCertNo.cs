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

namespace NetDataAccess.Extended.Jianzhu.QGJZSCJGGGFWPT
{
    public class GetFirstListPageByCertNo : ExternalRunWebPage
    {

        private int CompanyCountPerPage
        {
            get
            {
                return 15;
            }
        }
        
        public override byte[] GetRequestData_BeforeSendRequest(string pageUrl, Dictionary<string, string> listRow, Encoding encoding)
        {
            string certNo = listRow["certNo"];
            string data = "qy_type=&apt_scope=&apt_code=&qy_name=&qy_code=&apt_certno=" + certNo + "&qy_fr_name=&qy_gljg=&qy_reg_addr=&qy_region=";
            return encoding.GetBytes(data);
        }
        public override void WebRequestHtml_BeforeSendRequest(string pageUrl, Dictionary<string, string> listRow, Base.Web.NDAWebClient client)
        {
            client.Headers.Add("content-type", "application/x-www-form-urlencoded");
        }

        public override void CheckRequestCompleteFile(string webPageText, Dictionary<string, string> listRow)
        {
            if (webPageText.Contains("企业注册属地")  && webPageText.Trim().EndsWith("</html>"))
            {
                //完整获取了页面
            }
            else
            {
                throw new Exception("获取的页面不完整.");
            }
        }

        public override bool AfterAllGrab(IListSheet listSheet)
        {
            try
            {
                return this.GetCertNoCompCountList(listSheet) && this.GetListPageUrls(listSheet) && this.GetProvincePageUrlLists(listSheet);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private bool GetCertNoCompCountList(IListSheet listSheet)
        {

            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("certNo", 0); 
            resultColumnDic.Add("companyCount", 1);

            string resultFilePath = Path.Combine(exportDir, "各CertNo对应企业个数.xlsx");

            Dictionary<string, string> resultColumnFormat = new Dictionary<string, string>();
            resultColumnFormat.Add("companyCount", "#,##0");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, resultColumnFormat);

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string url = row[detailPageUrlColumnName];
                    string certNo = row["certNo"]; 

                    HtmlAgilityPack.HtmlDocument pageHtmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                    string pageText = pageHtmlDoc.DocumentNode.SelectSingleNode("//form[@class=\"pagingform\"]").NextSibling.NextSibling.InnerText;
                    int totalStartIndex = pageText.IndexOf("\"$total\":") + 9;
                    int totalEndIndex = pageText.IndexOf(",", totalStartIndex);
                    string totalCountStr = pageText.Substring(totalStartIndex, totalEndIndex - totalStartIndex);
                    int companyCount = int.Parse(totalCountStr);

                    Dictionary<string, object> f2vs = new Dictionary<string, object>();
                    f2vs.Add("certNo", certNo); 
                    f2vs.Add("companyCount", companyCount);
                    resultEW.AddRow(f2vs);
                }
            }

            resultEW.SaveToDisk();

            return true;
        }


        #region GetListPageUrls
        /// <summary>
        /// GetListPageUrls
        /// </summary>
        /// <param name="listSheet"></param>
        /// <param name="pageSourceDir"></param>
        /// <param name="resultEW"></param>
        private bool GetListPageUrls(IListSheet listSheet)
        { 
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
             
            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{ 
                "detailPageUrl",
                "detailPageName", 
                "cookie",
                "grabStatus", 
                "giveUpGrab", 
                "certNo"});

            string exportDir = this.RunPage.GetExportDir();

            string resultFilePath = Path.Combine(exportDir, "企业数据_证书编码查询企业列表页首页.xlsx");

            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic);
            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string url = row[detailPageUrlColumnName];
                    string certNo = row["certNo"];

                    HtmlAgilityPack.HtmlDocument pageHtmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                    string pageText = pageHtmlDoc.DocumentNode.SelectSingleNode("//form[@class=\"pagingform\"]").NextSibling.NextSibling.InnerText;
                    int totalStartIndex = pageText.IndexOf("\"$total\":") + 9;
                    int totalEndIndex = pageText.IndexOf(",", totalStartIndex);
                    string totalCountStr = pageText.Substring(totalStartIndex, totalEndIndex - totalStartIndex);
                    int companyCount = int.Parse(totalCountStr);
                    if (companyCount <= 450)
                    { 
                        Dictionary<string, string> f2vs = new Dictionary<string, string>();
                        f2vs.Add("detailPageUrl", "http://jzsc.mohurd.gov.cn/dataservice/query/comp/list?certNo=" + certNo);
                        f2vs.Add("detailPageName", certNo);
                        f2vs.Add("cookie", "filter_comp=show; JSESSIONID=DC4BC03F99DEDEBEFEE739B680BC5230; Hm_lvt_b1b4b9ea61b6f1627192160766a9c55c=1513578016,1513646440,1514281557,1514350446; Hm_lpvt_b1b4b9ea61b6f1627192160766a9c55c=1514356771");
                        f2vs.Add("certNo", certNo);
                        resultEW.AddRow(f2vs);
                    }
                    else
                    {
                        for (int j = 0; j < 10; j++)
                        {
                            string newCertNo = certNo + j.ToString();
                            Dictionary<string, string> f2vs = new Dictionary<string, string>();
                            f2vs.Add("detailPageUrl", "http://jzsc.mohurd.gov.cn/dataservice/query/comp/list?certNo=" + newCertNo);
                            f2vs.Add("detailPageName", newCertNo);
                            f2vs.Add("cookie", "filter_comp=show; JSESSIONID=DC4BC03F99DEDEBEFEE739B680BC5230; Hm_lvt_b1b4b9ea61b6f1627192160766a9c55c=1513578016,1513646440,1514281557,1514350446; Hm_lpvt_b1b4b9ea61b6f1627192160766a9c55c=1514356771");
                            f2vs.Add("certNo", newCertNo);
                            resultEW.AddRow(f2vs);
                        }

                    } 
                }
            }
 
            resultEW.SaveToDisk();
            return true;
        }
        #endregion

        private bool GetProvincePageUrlLists(IListSheet listSheet)
        {

            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("certNo", 5); 
            resultColumnDic.Add("pageIndex", 6);
            resultColumnDic.Add("perPageCount", 7);
            resultColumnDic.Add("companyCount", 8);
            string resultFilePath = Path.Combine(exportDir, "企业数据_各CertNo全部.xlsx");
            Dictionary<string, string> resultColumnFormat = new Dictionary<string, string>();
            resultColumnFormat.Add("pageIndex", "#,##0");
            resultColumnFormat.Add("perPageCount", "#,##0");
            resultColumnFormat.Add("companyCount", "#,##0");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, resultColumnFormat);

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string url = row[detailPageUrlColumnName];
                    string certNo = row["certNo"]; 
                    string cookie = row["cookie"];

                    HtmlAgilityPack.HtmlDocument pageHtmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                    string pageText = pageHtmlDoc.DocumentNode.SelectSingleNode("//form[@class=\"pagingform\"]").NextSibling.NextSibling.InnerText;
                    int totalStartIndex = pageText.IndexOf("\"$total\":") + 9;
                    int totalEndIndex = pageText.IndexOf(",", totalStartIndex);
                    string totalCountStr = pageText.Substring(totalStartIndex, totalEndIndex - totalStartIndex);
                    int companyCount = int.Parse(totalCountStr);

                    int pageIndex = 0;
                    while (pageIndex * this.CompanyCountPerPage < companyCount)
                    {
                        string detailPageUrl = "http://jzsc.mohurd.gov.cn/dataservice/query/comp/list?certNo=" + certNo + "&pageindex=" + pageIndex;
                        Dictionary<string, object> f2vs = new Dictionary<string, object>();
                        f2vs.Add("detailPageUrl", detailPageUrl);
                        f2vs.Add("detailPageName", certNo + "_" + pageIndex.ToString());
                        f2vs.Add("certNo", certNo); 
                        f2vs.Add("pageIndex", pageIndex + 1);
                        f2vs.Add("perPageCount", this.CompanyCountPerPage); 
                        f2vs.Add("companyCount", companyCount);
                        f2vs.Add("cookie", cookie);
                        resultEW.AddRow(f2vs);
                        pageIndex++;
                    }
                }
            }

            resultEW.SaveToDisk();

            return true;
        } 
    }
}