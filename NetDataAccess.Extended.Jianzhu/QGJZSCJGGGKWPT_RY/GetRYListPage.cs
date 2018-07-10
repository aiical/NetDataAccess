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

namespace NetDataAccess.Extended.Jianzhu.QGJZSCJGGGFWPT_RY
{
    public class GetRYListPage : ExternalRunWebPage
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
                return this.GetAllListPageUrls(listSheet);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private ExcelWriter GetExcelWriter(int pageIndex)
        {
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("formData", 5);

            string resultFilePath = Path.Combine(exportDir, "人员数据_人员列表页首页_" + pageIndex.ToString() + ".xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }

        private ExcelWriter GetAllExcelWriter(int allListFileIndex)
        {
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("formData", 5);
            resultColumnDic.Add("pageIndex", 6);
            resultColumnDic.Add("code", 7);

            string resultFilePath = Path.Combine(exportDir, "人员数据_人员所有列表页_" + allListFileIndex.ToString() + ".xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }



        private bool GetAllListPageUrls(IListSheet listSheet)
        {
            int pageIndex = 1;
            bool needMoreFirstPage = false;
            {
                ExcelWriter ew = null;

                string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;
                Dictionary<string, string> companyDic = new Dictionary<string, string>();
                for (int i = 0; i < listSheet.RowCount; i++)
                {
                    if (ew == null || ew.RowCount > 500000)
                    {
                        if (ew != null)
                        {
                            ew.SaveToDisk();
                        }
                        ew = this.GetExcelWriter(pageIndex);
                        pageIndex++;
                    }

                    Dictionary<string, string> row = listSheet.GetRow(i);
                    string detailPageUrl = row[SysConfig.DetailPageUrlFieldName];
                    string detailPageName = row[SysConfig.DetailPageNameFieldName];
                    string cookie = row[SysConfig.DetailPageCookieFieldName];

                    bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                    if (!giveUp)
                    {
                        HtmlAgilityPack.HtmlDocument pageHtmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);

                        HtmlNode pageNode = pageHtmlDoc.DocumentNode.SelectSingleNode("//a[@sf=\"pagebar\"]");
                        if (pageNode != null)
                        {
                            string pageData = pageNode.GetAttributeValue("sf:data", "");
                            if (pageData.Length == 0)
                            {
                                throw new Exception("获取分页信息错误. detailPageName = " + detailPageName);
                            }
                            else
                            {
                                JObject rootJo = JObject.Parse(pageData.Substring(1, pageData.Length - 2));
                                string ps = rootJo.GetValue("ps").ToString();
                                string tt = rootJo.GetValue("tt").ToString();
                                string pc = rootJo.GetValue("pc").ToString();
                                int pageCount = int.Parse(pc);
                                if (pageCount < 30)
                                {
                                    string formDataCode = detailPageName.Replace("%", "%25");
                                    Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                    f2vs.Add("detailPageUrl", "http://jzsc.mohurd.gov.cn/dataservice/query/staff/list?_=" + detailPageName);
                                    f2vs.Add("detailPageName", detailPageName);
                                    f2vs.Add("cookie", cookie);
                                    f2vs.Add("formData", "ry_type=&ry_reg_type=&ry_name=&reg_seal_code=&ry_cardno=&ry_qymc=&complexname=" + formDataCode);
                                    ew.AddRow(f2vs);
                                }
                                else
                                {
                                    needMoreFirstPage = true;
                                    if (detailPageName.Length == 9)
                                    {
                                        for (int j = 0; j < 10; j++)
                                        {
                                            string code = detailPageName.Insert(2, j.ToString());
                                            string formDataCode = code.Replace("%", "%25");
                                            Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                            f2vs.Add("detailPageUrl", "http://jzsc.mohurd.gov.cn/dataservice/query/staff/list?_=" + code);
                                            f2vs.Add("detailPageName", code);
                                            f2vs.Add("cookie", cookie);
                                            f2vs.Add("formData", "ry_type=&ry_reg_type=&ry_name=&reg_seal_code=&ry_cardno=&ry_qymc=&complexname=" + formDataCode);
                                            ew.AddRow(f2vs);
                                        }
                                    }
                                    else if (detailPageName.Length == 10)
                                    {
                                        for (int j = 0; j < 10; j++)
                                        {
                                            string code = detailPageName.Insert(3, j.ToString());
                                            string formDataCode = code.Replace("%", "%25");
                                            Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                            f2vs.Add("detailPageUrl", "http://jzsc.mohurd.gov.cn/dataservice/query/staff/list?_=" + code);
                                            f2vs.Add("detailPageName", code);
                                            f2vs.Add("cookie", cookie);
                                            f2vs.Add("formData", "ry_type=&ry_reg_type=&ry_name=&reg_seal_code=&ry_cardno=&ry_qymc=&complexname=" + formDataCode);
                                            ew.AddRow(f2vs);
                                        }
                                    }
                                    else if (detailPageName.Length == 11)
                                    {
                                        for (int j = 0; j < 10; j++)
                                        {
                                            string code = detailPageName.Insert(4, j.ToString());
                                            string formDataCode = code.Replace("%", "%25");
                                            Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                            f2vs.Add("detailPageUrl", "http://jzsc.mohurd.gov.cn/dataservice/query/staff/list?_=" + code);
                                            f2vs.Add("detailPageName", code);
                                            f2vs.Add("cookie", cookie);
                                            f2vs.Add("formData", "ry_type=&ry_reg_type=&ry_name=&reg_seal_code=&ry_cardno=&ry_qymc=&complexname=" + formDataCode);
                                            ew.AddRow(f2vs);
                                        }
                                    }
                                    else if (detailPageName.Length == 12)
                                    {
                                        for (int j = 0; j < 10; j++)
                                        {
                                            string code = detailPageName.Insert(5, j.ToString());
                                            string formDataCode = code.Replace("%", "%25");
                                            Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                            f2vs.Add("detailPageUrl", "http://jzsc.mohurd.gov.cn/dataservice/query/staff/list?_=" + code);
                                            f2vs.Add("detailPageName", code);
                                            f2vs.Add("cookie", cookie);
                                            f2vs.Add("formData", "ry_type=&ry_reg_type=&ry_name=&reg_seal_code=&ry_cardno=&ry_qymc=&complexname=" + formDataCode);
                                            ew.AddRow(f2vs);
                                        }
                                    }
                                    else
                                    {
                                        for (int j = 1; j <= 12; j++)
                                        {
                                            int dayCount = 0;
                                            switch (j)
                                            {
                                                case 1:
                                                case 3:
                                                case 5:
                                                case 7:
                                                case 8:
                                                case 10:
                                                case 12:
                                                    dayCount = 31;
                                                    break;
                                                case 2:
                                                    dayCount = 29;
                                                    break;
                                                case 4:
                                                case 6:
                                                case 9:
                                                case 11:
                                                    dayCount = 30;
                                                    break;
                                            }
                                            for (int k = 1; k <= dayCount; k++)
                                            {
                                                for (int l = 0; l < 100; l++)
                                                {
                                                    string code = detailPageName + "%" + j.ToString().PadLeft(2, '0') + k.ToString().PadLeft(2, '0') + l.ToString().PadLeft(2, '0');
                                                    string formDataCode = code.Replace("%", "%25");
                                                    Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                                    f2vs.Add("detailPageUrl", "http://jzsc.mohurd.gov.cn/dataservice/query/staff/list?_=" + code);
                                                    f2vs.Add("detailPageName", code);
                                                    f2vs.Add("cookie", cookie);
                                                    f2vs.Add("formData", "ry_type=&ry_reg_type=&ry_name=&reg_seal_code=&ry_cardno=&ry_qymc=&complexname=" + formDataCode);
                                                    ew.AddRow(f2vs);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                ew.SaveToDisk();
            }

            if (!needMoreFirstPage)
            {
                int allListFileIndex = 1;
                ExcelWriter ew = null;

                string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;
                Dictionary<string, string> companyDic = new Dictionary<string, string>();
                for (int i = 0; i < listSheet.RowCount; i++)
                {
                    if (ew == null || ew.RowCount > 500000)
                    {
                        if (ew != null)
                        {
                            ew.SaveToDisk();
                        }
                        ew = this.GetAllExcelWriter(allListFileIndex);
                        allListFileIndex++;
                    }
                    Dictionary<string, string> row = listSheet.GetRow(i);
                    string detailPageUrl = row[SysConfig.DetailPageUrlFieldName];
                    string detailPageName = row[SysConfig.DetailPageNameFieldName];
                    string cookie = row[SysConfig.DetailPageCookieFieldName];

                    bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                    if (!giveUp)
                    {
                        HtmlAgilityPack.HtmlDocument pageHtmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);

                        HtmlNode pageNode = pageHtmlDoc.DocumentNode.SelectSingleNode("//a[@sf=\"pagebar\"]");
                        if (pageNode != null)
                        {
                            string pageData = pageNode.GetAttributeValue("sf:data", "");
                            if (pageData.Length == 0)
                            {
                                throw new Exception("获取分页信息错误. detailPageName = " + detailPageName);
                            }
                            else
                            {
                                Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                f2vs.Add("detailPageUrl", "http://jzsc.mohurd.gov.cn/dataservice/query/project/list?_=" + detailPageName);
                                f2vs.Add("detailPageName", detailPageName+ "_1");
                                f2vs.Add("cookie", "filter_comp=; JSESSIONID=F1DC2E6DC10B3E64CC59C070A5722639; Hm_lvt_b1b4b9ea61b6f1627192160766a9c55c=1515293273,1515384893,1515553333,1515638274; Hm_lpvt_b1b4b9ea61b6f1627192160766a9c55c=1515645794");
                                f2vs.Add("formData", "jsxm_name=&cons_name=&jsxm_region=&jsxm_region_id=&complexname=" + detailPageName);
                                f2vs.Add("pageIndex", "1");
                                f2vs.Add("code", detailPageName);
                                ew.AddRow(f2vs);

                                JObject rootJo = JObject.Parse(pageData.Substring(1, pageData.Length - 2));
                                string ps = rootJo.GetValue("ps").ToString();
                                string tt = rootJo.GetValue("tt").ToString();
                                string pc = rootJo.GetValue("pc").ToString();
                                int pageCount = int.Parse(pc);
                                for (int pIndex = 2; pIndex <= pageCount; pIndex++)
                                {
                                    Dictionary<string, string> otherF2vs = new Dictionary<string, string>();
                                    otherF2vs.Add("detailPageUrl", "http://jzsc.mohurd.gov.cn/dataservice/query/project/list?_=" + detailPageName + "_" + pIndex);
                                    otherF2vs.Add("detailPageName", detailPageName + "_" + pIndex.ToString());
                                    otherF2vs.Add("cookie", "filter_comp=; JSESSIONID=F1DC2E6DC10B3E64CC59C070A5722639; Hm_lvt_b1b4b9ea61b6f1627192160766a9c55c=1515293273,1515384893,1515553333,1515638274; Hm_lpvt_b1b4b9ea61b6f1627192160766a9c55c=1515645794");
                                    otherF2vs.Add("formData", "complexname=" + detailPageName + "&jsxm_region_id=&%24total=" + tt + "&%24reload=0&jsxm_region=&jsxm_name=&cons_name=&%24pg=" + pIndex.ToString() + "&%24pgsz=15");
                                    otherF2vs.Add("code", detailPageName);
                                    otherF2vs.Add("pageIndex", pIndex.ToString());
                                    ew.AddRow(otherF2vs);
                                }
                            }
                        }
                    }
                }

                ew.SaveToDisk();
            }

            return true;
        }
    }
}