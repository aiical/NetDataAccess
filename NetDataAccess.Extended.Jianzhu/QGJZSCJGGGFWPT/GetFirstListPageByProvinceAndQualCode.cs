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
    public class GetFirstListPageByProvinceAndQualCode : ExternalRunWebPage
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
            string regionName = listRow["regionFullName"];
            string regionId = listRow["regionId"];
            string aptCode = listRow["aptCode"];
            string aptScope = listRow["aptScope"];
            string encodeRegionName = System.Web.HttpUtility.UrlEncode(regionName);
            string encodeAptScope = System.Web.HttpUtility.UrlEncode(aptScope);
            string data = "qy_type=&apt_scope=" + aptScope + "&apt_code=" + aptCode + "&qy_name=&qy_code=&apt_certno=&qy_fr_name=&qy_gljg=&qy_reg_addr=" + encodeRegionName + "&qy_region=" + regionId;
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
                return this.GetProvinceCompCountList(listSheet) && this.GetProvincePageUrlLists(listSheet);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private bool GetProvinceCompCountList(IListSheet listSheet)
        {

            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("regionId", 0);
            resultColumnDic.Add("regionName", 1);
            resultColumnDic.Add("regionFullName", 2);
            resultColumnDic.Add("aptCode", 3);
            resultColumnDic.Add("aptScope", 4);
            resultColumnDic.Add("companyCount", 5);

            string resultFilePath = Path.Combine(exportDir, "各省企业个数.xlsx");

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
                    string provinceId = row["regionId"];
                    string provinceName = row["regionName"];
                    string provinceFullName = row["regionFullName"];
                    string aptCode = row["aptCode"];
                    string aptScope = row["aptScope"];
                    string cookie = row["cookie"];

                    HtmlAgilityPack.HtmlDocument pageHtmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                    string pageText = pageHtmlDoc.DocumentNode.SelectSingleNode("//form[@class=\"pagingform\"]").NextSibling.NextSibling.InnerText;
                    int totalStartIndex = pageText.IndexOf("\"$total\":") + 9;
                    int totalEndIndex = pageText.IndexOf(",", totalStartIndex);
                    string totalCountStr = pageText.Substring(totalStartIndex, totalEndIndex - totalStartIndex);
                    int companyCount = int.Parse(totalCountStr);

                    Dictionary<string, object> f2vs = new Dictionary<string, object>();
                    f2vs.Add("regionId", provinceId);
                    f2vs.Add("regionName", provinceName);
                    f2vs.Add("regionFullName", provinceFullName);
                    f2vs.Add("aptCode", aptCode);
                    f2vs.Add("aptScope", aptScope);
                    f2vs.Add("companyCount", companyCount);
                    resultEW.AddRow(f2vs);
                }
            }

            resultEW.SaveToDisk();

            return true;
        }
        private bool GetProvincePageUrlLists(IListSheet listSheet)
        {

            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("regionId", 5);
            resultColumnDic.Add("regionName", 6);
            resultColumnDic.Add("regionFullName", 7);
            resultColumnDic.Add("aptCode", 8);
            resultColumnDic.Add("aptScope", 9);
            resultColumnDic.Add("pageIndex", 10);
            resultColumnDic.Add("perPageCount", 11);
            resultColumnDic.Add("companyCount", 12);
            string resultFilePath = Path.Combine(exportDir, "企业数据_各省份全部.xlsx");
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
                    string provinceId = row["regionId"];
                    string provinceName = row["regionName"];
                    string provinceFullName = row["regionFullName"];
                    string aptCode = row["aptCode"];
                    string aptScope = row["aptScope"];
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
                        string detailPageUrl = "http://jzsc.mohurd.gov.cn/dataservice/query/comp/list?regionname=" + provinceName + "&aptScope=" + aptScope + "&pageindex=" + pageIndex;
                        Dictionary<string, object> f2vs = new Dictionary<string, object>();
                        f2vs.Add("detailPageUrl", detailPageUrl);
                        f2vs.Add("detailPageName", provinceId + "_" + aptScope + "_" + pageIndex.ToString());
                        f2vs.Add("regionId", provinceId);
                        f2vs.Add("regionName", provinceName);
                        f2vs.Add("regionFullName", provinceFullName);
                        f2vs.Add("pageIndex", pageIndex + 1);
                        f2vs.Add("perPageCount", this.CompanyCountPerPage);
                        f2vs.Add("aptCode", aptCode);
                        f2vs.Add("aptScope", aptScope);
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