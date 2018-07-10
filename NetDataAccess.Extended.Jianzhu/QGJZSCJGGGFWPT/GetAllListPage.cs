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
    public class GetAllListPage : ExternalRunWebPage
    {
        public override byte[] GetRequestData_BeforeSendRequest(string pageUrl, Dictionary<string, string> listRow, Encoding encoding)
        {
            string certNo = listRow["certNo"];
            string pageIndex = listRow["pageIndex"];
            string perPageCount = listRow["perPageCount"];
            string companyCount = listRow["companyCount"];
            string data = "apt_code=&qy_fr_name=&%24total=" + companyCount + "&qy_reg_addr=&qy_code=&qy_name=&%24pgsz=" + perPageCount.ToString() + "&apt_certno=" + certNo + "&qy_region=&%24reload=0&qy_type=&%24pg=" + pageIndex.ToString() + "&qy_gljg=&apt_scope=";
            return encoding.GetBytes(data);
        } 
        public override void WebRequestHtml_BeforeSendRequest(string pageUrl, Dictionary<string, string> listRow, Base.Web.NDAWebClient client)
        {
            client.Headers.Add("content-type", "application/x-www-form-urlencoded");
        }

        public override bool AfterAllGrab(IListSheet listSheet)
        {
            try
            {
                return //this.GetQYGCXMPageUrls(listSheet)
                    //&& this.GetCompanyInfoPageUrls(listSheet)
                    //&& this.GetQYZZZGPageUrls(listSheet)
                    //&& 
                    this.GetQYZCRYSYPageUrls(listSheet)
                    //&& this.GetQYBLJLPageUrls(listSheet)
                    //&& this.GetQYLHJLPageUrls(listSheet)
                    //&& this.GetQYBGJLPageUrls(listSheet)
                    ; 
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private ExcelWriter GetCompanyInfoExcelWriter()
        {
            //企业基本信息
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4); 
            resultColumnDic.Add("companyId", 5);

            string resultFilePath = Path.Combine(exportDir, "企业数据_企业信息.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }

        private ExcelWriter GetQYZZZGExcelWriter()
        {
            //企业资质资格
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("companyId", 5);
            string resultFilePath = Path.Combine(exportDir, "企业数据_企业资质资格首页.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }

        private ExcelWriter GetQYZCRYSYExcelWriter()
        {
            //企业注册人员首页
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("companyId", 5);
            string resultFilePath = Path.Combine(exportDir, "企业数据_企业注册人员首页.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }

        private ExcelWriter GetQYGCXMExcelWriter()
        {
            //企业工程项目
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("companyId", 5);
            string resultFilePath = Path.Combine(exportDir, "企业数据_企业工程项目首页.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }

        private ExcelWriter GetQYBLJLExcelWriter()
        {
            //企业不良记录
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("companyId", 5);
            string resultFilePath = Path.Combine(exportDir, "企业数据_企业不良记录首页.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }

        private ExcelWriter GetQYLHJLExcelWriter()
        {
            //企业良好记录
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("companyId", 5);
            string resultFilePath = Path.Combine(exportDir, "企业数据_企业良好记录首页.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }

        private ExcelWriter GetQYBGJLExcelWriter()
        {
            //企业变更记录
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("companyId", 5);
            string resultFilePath = Path.Combine(exportDir, "企业数据_企业变更记录首页.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }

        private bool GetCompanyInfoPageUrls(IListSheet listSheet)
        {
            ExcelWriter companyInfoEW = this.GetCompanyInfoExcelWriter();

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;
            Dictionary<string, string> companyDic = new Dictionary<string, string>();

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    HtmlAgilityPack.HtmlDocument pageHtmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                    HtmlNodeCollection allCompanyNodes = pageHtmlDoc.DocumentNode.SelectNodes("//div[@class=\"mtop\"]/table/tbody/tr");
                    if (allCompanyNodes != null)
                    {
                        foreach (HtmlNode companyNode in allCompanyNodes)
                        {
                            try
                            {
                                HtmlNode companyLinkNode = companyNode.SelectSingleNode("./td/a");
                                string companyId = "";
                                if (companyLinkNode != null)
                                {
                                    string companyUrl = companyLinkNode.GetAttributeValue("href", "");
                                    int companyIdStartIndex = companyUrl.LastIndexOf("/") + 1;
                                    companyId = companyUrl.Substring(companyIdStartIndex);
                                }
                                if (companyId.Length > 0 && !companyDic.ContainsKey(companyId))
                                {
                                    companyDic.Add(companyId, null);
                                    Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                    f2vs.Add("detailPageUrl", "http://jzsc.mohurd.gov.cn/dataservice/query/comp/compDetail/" + companyId);
                                    f2vs.Add("detailPageName", companyId);
                                    f2vs.Add("companyId", companyId);
                                    companyInfoEW.AddRow(f2vs);
                                }
                            }
                            catch (Exception ex)
                            {
                                throw ex;
                            }
                        }
                    }
                }
            }

            companyInfoEW.SaveToDisk();

            return true;
        }

        private bool GetQYZZZGPageUrls(IListSheet listSheet)
        { 
            ExcelWriter qyzzzgEW = this.GetQYZZZGExcelWriter(); 

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;
            Dictionary<string, string> companyDic = new Dictionary<string, string>();

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    HtmlAgilityPack.HtmlDocument pageHtmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                    HtmlNodeCollection allCompanyNodes = pageHtmlDoc.DocumentNode.SelectNodes("//div[@class=\"mtop\"]/table/tbody/tr");
                    if (allCompanyNodes != null)
                    {
                        foreach (HtmlNode companyNode in allCompanyNodes)
                        {
                            try
                            {
                                HtmlNode companyLinkNode = companyNode.SelectSingleNode("./td/a");
                                string companyId = "";
                                if (companyLinkNode != null)
                                {
                                    string companyUrl = companyLinkNode.GetAttributeValue("href", "");
                                    int companyIdStartIndex = companyUrl.LastIndexOf("/") + 1;
                                    companyId = companyUrl.Substring(companyIdStartIndex);
                                }
                                if (companyId.Length > 0 && !companyDic.ContainsKey(companyId))
                                {
                                    companyDic.Add(companyId, null);
                                    Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                    f2vs.Add("detailPageUrl", "http://jzsc.mohurd.gov.cn/dataservice/query/comp/caDetailList/" + companyId + "?_=1513660572729");
                                    f2vs.Add("detailPageName", companyId);
                                    f2vs.Add("companyId", companyId);
                                    qyzzzgEW.AddRow(f2vs);
                                }
                            }
                            catch (Exception ex)
                            {
                                throw ex;
                            }
                        }
                    }
                }
            }
             
            qyzzzgEW.SaveToDisk(); 

            return true;
        }

        private bool GetQYZCRYSYPageUrls(IListSheet listSheet)
        { 
            ExcelWriter qyzcrysyEW = this.GetQYZCRYSYExcelWriter(); 

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;
            Dictionary<string, string> companyDic = new Dictionary<string, string>();

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    HtmlAgilityPack.HtmlDocument pageHtmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                    HtmlNodeCollection allCompanyNodes = pageHtmlDoc.DocumentNode.SelectNodes("//div[@class=\"mtop\"]/table/tbody/tr");
                    if (allCompanyNodes != null)
                    {
                        foreach (HtmlNode companyNode in allCompanyNodes)
                        {
                            try
                            {
                                HtmlNode companyLinkNode = companyNode.SelectSingleNode("./td/a");
                                string companyId = "";
                                if (companyLinkNode != null)
                                {
                                    string companyUrl = companyLinkNode.GetAttributeValue("href", "");
                                    int companyIdStartIndex = companyUrl.LastIndexOf("/") + 1;
                                    companyId = companyUrl.Substring(companyIdStartIndex);
                                }
                                if (companyId.Length > 0 && !companyDic.ContainsKey(companyId))
                                {
                                    companyDic.Add(companyId, null);
                                    Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                    f2vs.Add("detailPageUrl", "http://jzsc.mohurd.gov.cn/dataservice/query/comp/regStaffList/" + companyId);
                                    f2vs.Add("detailPageName", companyId);
                                    f2vs.Add("companyId", companyId);
                                    qyzcrysyEW.AddRow(f2vs);
                                }
                            }
                            catch (Exception ex)
                            {
                                throw ex;
                            }
                        }
                    }
                }
            } 

            qyzcrysyEW.SaveToDisk(); 

            return true;
        }

        private bool GetQYBLJLPageUrls(IListSheet listSheet)
        {
            ExcelWriter qybljlEW = this.GetQYBLJLExcelWriter();


            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;
            Dictionary<string, string> companyDic = new Dictionary<string, string>();

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    HtmlAgilityPack.HtmlDocument pageHtmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                    HtmlNodeCollection allCompanyNodes = pageHtmlDoc.DocumentNode.SelectNodes("//div[@class=\"mtop\"]/table/tbody/tr");
                    if (allCompanyNodes != null)
                    {
                        foreach (HtmlNode companyNode in allCompanyNodes)
                        {
                            try
                            {
                                HtmlNode companyLinkNode = companyNode.SelectSingleNode("./td/a");
                                string companyId = "";
                                if (companyLinkNode != null)
                                {
                                    string companyUrl = companyLinkNode.GetAttributeValue("href", "");
                                    int companyIdStartIndex = companyUrl.LastIndexOf("/") + 1;
                                    companyId = companyUrl.Substring(companyIdStartIndex);
                                }
                                if (companyId.Length > 0 && !companyDic.ContainsKey(companyId))
                                {
                                    companyDic.Add(companyId, null);
                                    Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                    f2vs.Add("detailPageUrl", "http://jzsc.mohurd.gov.cn/dataservice/query/comp/compCreditRecordList/" + companyId + "/0");
                                    f2vs.Add("detailPageName", companyId);
                                    f2vs.Add("companyId", companyId);
                                    qybljlEW.AddRow(f2vs);
                                }
                            }
                            catch (Exception ex)
                            {
                                throw ex;
                            }
                        }
                    }
                }
            }

            qybljlEW.SaveToDisk();
            return true;
        }

        private bool GetQYGCXMPageUrls(IListSheet listSheet)
        {
            ExcelWriter qygcxmEW = this.GetQYGCXMExcelWriter();


            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;
            Dictionary<string, string> companyDic = new Dictionary<string, string>();

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    HtmlAgilityPack.HtmlDocument pageHtmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                    HtmlNodeCollection allCompanyNodes = pageHtmlDoc.DocumentNode.SelectNodes("//div[@class=\"mtop\"]/table/tbody/tr");
                    if (allCompanyNodes != null)
                    {
                        foreach (HtmlNode companyNode in allCompanyNodes)
                        {
                            try
                            {
                                HtmlNode companyLinkNode = companyNode.SelectSingleNode("./td/a");
                                string companyId = "";
                                if (companyLinkNode != null)
                                {
                                    string companyUrl = companyLinkNode.GetAttributeValue("href", "");
                                    int companyIdStartIndex = companyUrl.LastIndexOf("/") + 1;
                                    companyId = companyUrl.Substring(companyIdStartIndex);
                                }
                                if (companyId.Length > 0 && !companyDic.ContainsKey(companyId))
                                {
                                    companyDic.Add(companyId, null);
                                    Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                    f2vs.Add("detailPageUrl", "http://jzsc.mohurd.gov.cn/dataservice/query/comp/compPerformanceListSys/" + companyId);
                                    f2vs.Add("detailPageName", companyId);
                                    f2vs.Add("companyId", companyId);
                                    qygcxmEW.AddRow(f2vs);
                                }
                            }
                            catch (Exception ex)
                            {
                                throw ex;
                            }
                        }
                    }
                }
            }

            qygcxmEW.SaveToDisk();
            return true;
        }

        private bool GetQYLHJLPageUrls(IListSheet listSheet)
        { 
            ExcelWriter qylhjlEW = this.GetQYLHJLExcelWriter(); 

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;
            Dictionary<string, string> companyDic = new Dictionary<string, string>();

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    HtmlAgilityPack.HtmlDocument pageHtmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                    HtmlNodeCollection allCompanyNodes = pageHtmlDoc.DocumentNode.SelectNodes("//div[@class=\"mtop\"]/table/tbody/tr");
                    if (allCompanyNodes != null)
                    {
                        foreach (HtmlNode companyNode in allCompanyNodes)
                        {
                            try
                            {
                                HtmlNode companyLinkNode = companyNode.SelectSingleNode("./td/a");
                                string companyId = "";
                                if (companyLinkNode != null)
                                {
                                    string companyUrl = companyLinkNode.GetAttributeValue("href", "");
                                    int companyIdStartIndex = companyUrl.LastIndexOf("/") + 1;
                                    companyId = companyUrl.Substring(companyIdStartIndex);
                                }
                                if (companyId.Length > 0 && !companyDic.ContainsKey(companyId))
                                {
                                    companyDic.Add(companyId, null);
                                    Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                    f2vs.Add("detailPageUrl", "http://jzsc.mohurd.gov.cn/dataservice/query/comp/compCreditRecordList/" + companyId + "/1");
                                    f2vs.Add("detailPageName", companyId);
                                    f2vs.Add("companyId", companyId);
                                    qylhjlEW.AddRow(f2vs);
                                }
                            }
                            catch (Exception ex)
                            {
                                throw ex;
                            }
                        }
                    }
                }
            }
             
            qylhjlEW.SaveToDisk(); 

            return true;
        }

        private bool GetQYBGJLPageUrls(IListSheet listSheet)
        { 
            ExcelWriter qybgjlEW = this.GetQYBGJLExcelWriter();

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;
            Dictionary<string, string> companyDic = new Dictionary<string, string>();

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    HtmlAgilityPack.HtmlDocument pageHtmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                    HtmlNodeCollection allCompanyNodes = pageHtmlDoc.DocumentNode.SelectNodes("//div[@class=\"mtop\"]/table/tbody/tr");
                    if (allCompanyNodes != null)
                    {
                        foreach (HtmlNode companyNode in allCompanyNodes)
                        {
                            try
                            {
                                HtmlNode companyLinkNode = companyNode.SelectSingleNode("./td/a");
                                string companyId = "";
                                if (companyLinkNode != null)
                                {
                                    string companyUrl = companyLinkNode.GetAttributeValue("href", "");
                                    int companyIdStartIndex = companyUrl.LastIndexOf("/") + 1;
                                    companyId = companyUrl.Substring(companyIdStartIndex);
                                }
                                if (companyId.Length > 0 && !companyDic.ContainsKey(companyId))
                                {
                                    companyDic.Add(companyId, null);
                                    Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                    f2vs.Add("detailPageUrl", "http://jzsc.mohurd.gov.cn/dataservice/query/comp/traceList/" + companyId);
                                    f2vs.Add("detailPageName", companyId);
                                    f2vs.Add("companyId", companyId);
                                    qybgjlEW.AddRow(f2vs);
                                }
                            }
                            catch (Exception ex)
                            {
                                throw ex;
                            }
                        }
                    }
                }
            }
            
            qybgjlEW.SaveToDisk();

            return true;
        } 
    }
}