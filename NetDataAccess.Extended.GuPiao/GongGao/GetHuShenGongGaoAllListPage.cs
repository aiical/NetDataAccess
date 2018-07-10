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
using System.Globalization;

namespace NetDataAccess.Extended.GuPiao
{
    public class GetHuShenGongGaoAllListPage : ExternalRunWebPage
    {
        public override byte[] GetRequestData_BeforeSendRequest(string pageUrl, Dictionary<string, string> listRow, Encoding encoding)
        {
            string data = listRow["formData"];
            return encoding.GetBytes(data);
        }

        public override void WebRequestHtml_BeforeSendRequest(string pageUrl, Dictionary<string, string> listRow, Base.Web.NDAWebClient client)
        {
            client.Headers.Add("content-type", "application/x-www-form-urlencoded");
        }

        public override bool AfterAllGrab(IListSheet listSheet)
        {
            this.GetGongGaoAllDetailPageUrls(listSheet);
            this.GetGongGaoAllDetailTxtFileLocalUrls(listSheet);
            this.GetGongGaoListAllPagesToCsv(listSheet);
            return true;
        }

        private ExcelWriter GetDetailPageExcelWriter(int fileIndex)
        {
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("pinyin", 5);
            resultColumnDic.Add("zwjc", 6);
            resultColumnDic.Add("code", 7);
            resultColumnDic.Add("orgId", 8);
            resultColumnDic.Add("stockExchange", 9);
            resultColumnDic.Add("category", 10);
            resultColumnDic.Add("announcementTitle", 11);
            resultColumnDic.Add("announcementTime", 12);
            resultColumnDic.Add("adjunctType", 13);
            string resultFilePath = Path.Combine(exportDir, "沪深股票公告内容页_" + fileIndex.ToString() + ".xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }

        private void GetGongGaoAllDetailPageUrls(IListSheet listSheet)
        {
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            int fileIndex = 1;
            ExcelWriter ew = null;
            Dictionary<string, string> announcementDic = new Dictionary<string, string>();

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                if (ew == null || ew.RowCount > 500000)
                {
                    if (ew != null)
                    {
                        ew.SaveToDisk();
                    }
                    ew = this.GetDetailPageExcelWriter(fileIndex);
                    fileIndex++;
                }

                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"];
                string orgId = row["orgId"];
                string pinyin = row["pinyin"];
                string code = row["code"];
                string zwjc = row["zwjc"];
                string category = row["category"];
                string stockExchange = row["stockExchange"];

                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string localFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir);
                    TextReader tr = null;

                    try
                    {
                        tr = new StreamReader(localFilePath, Encoding.UTF8);
                        string js = tr.ReadToEnd();

                        JObject rootJo = JObject.Parse(js);
                        JArray itemArrayJsons = rootJo.SelectToken("announcements") as JArray;
                        for (int j = 0; j < itemArrayJsons.Count; j++)
                        {
                            JObject itemJson = itemArrayJsons[j] as JObject;
                            string announcementId = itemJson.GetValue("announcementId").ToString().Trim();
                            string announcementTitle = CommonUtil.HtmlDecode(itemJson.GetValue("announcementTitle").ToString().Trim());
                            string announcementTimeStr = itemJson.GetValue("announcementTime").ToString().Trim();
                            string adjunctType = itemJson.GetValue("adjunctType").ToString().Trim();
                            string adjunctUrl = itemJson.GetValue("adjunctUrl").ToString().Trim();

                            DateTime announcementTime = (new DateTime(1970, 1, 1)).AddMilliseconds(long.Parse(announcementTimeStr)).ToLocalTime();
                            string outAnnouncementTimeStr = announcementTime.ToString("yyyy-MM-dd HH:mm:ss");

                            if (!announcementDic.ContainsKey(announcementId))
                            {
                                announcementDic.Add(announcementId, null);

                                Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                f2vs.Add("detailPageUrl", "http://www.cninfo.com.cn/" + adjunctUrl);
                                f2vs.Add("detailPageName", announcementId);
                                f2vs.Add("zwjc", zwjc);
                                f2vs.Add("code", code);
                                f2vs.Add("pinyin", pinyin);
                                f2vs.Add("orgId", orgId);
                                f2vs.Add("category", category);
                                f2vs.Add("stockExchange", stockExchange);
                                f2vs.Add("announcementTitle", announcementTitle);
                                f2vs.Add("announcementTime", outAnnouncementTimeStr);
                                f2vs.Add("adjunctType", adjunctType);
                                ew.AddRow(f2vs);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                    finally
                    {
                        if (tr != null)
                        {
                            tr.Close();
                            tr.Dispose();
                        }
                    }
                }
            }
            ew.SaveToDisk();
        }

        private ExcelWriter GetDetailTxtFileLocalUrlExcelWriter(int fileIndex)
        {
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("pinyin", 5);
            resultColumnDic.Add("zwjc", 6);
            resultColumnDic.Add("code", 7);
            resultColumnDic.Add("orgId", 8);
            resultColumnDic.Add("stockExchange", 9);
            resultColumnDic.Add("category", 10);
            resultColumnDic.Add("announcementTitle", 11);
            resultColumnDic.Add("announcementTime", 12);
            resultColumnDic.Add("adjunctType", 13);
            string resultFilePath = Path.Combine(exportDir, "沪深股票公告内容页TXT文件_" + fileIndex.ToString() + ".xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }

        private void GetGongGaoAllDetailTxtFileLocalUrls(IListSheet listSheet)
        {
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            int fileIndex = 1;
            ExcelWriter ew = null;
            Dictionary<string, string> announcementDic = new Dictionary<string, string>();

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                if (ew == null || ew.RowCount > 500000)
                {
                    if (ew != null)
                    {
                        ew.SaveToDisk();
                    }
                    ew = this.GetDetailTxtFileLocalUrlExcelWriter(fileIndex);
                    fileIndex++;
                }

                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"];
                string orgId = row["orgId"];
                string pinyin = row["pinyin"];
                string code = row["code"];
                string zwjc = row["zwjc"];
                string category = row["category"];
                string stockExchange = row["stockExchange"];

                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string localFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir);
                    TextReader tr = null;

                    try
                    {
                        tr = new StreamReader(localFilePath, Encoding.UTF8);
                        string js = tr.ReadToEnd();

                        JObject rootJo = JObject.Parse(js);
                        JArray itemArrayJsons = rootJo.SelectToken("announcements") as JArray;
                        for (int j = 0; j < itemArrayJsons.Count; j++)
                        {
                            JObject itemJson = itemArrayJsons[j] as JObject;
                            string announcementId = itemJson.GetValue("announcementId").ToString().Trim();
                            string announcementTitle = CommonUtil.HtmlDecode(itemJson.GetValue("announcementTitle").ToString().Trim());
                            string announcementTimeStr = itemJson.GetValue("announcementTime").ToString().Trim();
                            string adjunctType = itemJson.GetValue("adjunctType").ToString().Trim();
                            string adjunctUrl = itemJson.GetValue("adjunctUrl").ToString().Trim();

                            DateTime announcementTime = (new DateTime(1970, 1, 1)).AddMilliseconds(long.Parse(announcementTimeStr)).ToLocalTime();
                            string outAnnouncementTimeStr = announcementTime.ToString("yyyy-MM-dd HH:mm:ss");

                            if (!announcementDic.ContainsKey(announcementId))
                            {
                                announcementDic.Add(announcementId, null);

                                Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                f2vs.Add("detailPageUrl", "http://www.cninfo.com.cn/" + adjunctUrl);
                                f2vs.Add("detailPageName", announcementId);
                                f2vs.Add("zwjc", zwjc);
                                f2vs.Add("code", code);
                                f2vs.Add("pinyin", pinyin);
                                f2vs.Add("orgId", orgId);
                                f2vs.Add("category", category);
                                f2vs.Add("stockExchange", stockExchange);
                                f2vs.Add("announcementTitle", announcementTitle);
                                f2vs.Add("announcementTime", outAnnouncementTimeStr);
                                f2vs.Add("adjunctType", adjunctType);
                                ew.AddRow(f2vs);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                    finally
                    {
                        if (tr != null)
                        {
                            tr.Close();
                            tr.Dispose();
                        }
                    }
                }
            }
            ew.SaveToDisk();
        }

        private CsvWriter GetCsvWriter()
        {
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("pinyin", 0);
            resultColumnDic.Add("zwjc", 1);
            resultColumnDic.Add("code", 2);
            resultColumnDic.Add("orgId", 3);
            resultColumnDic.Add("stockExchange", 4);
            resultColumnDic.Add("category", 5);
            resultColumnDic.Add("announcementTitle", 6);
            resultColumnDic.Add("announcementTime", 7);
            resultColumnDic.Add("announcementId", 8);
            resultColumnDic.Add("adjunctType", 9);
            resultColumnDic.Add("fileUrl", 10);
            string resultFilePath = Path.Combine(exportDir, "沪深股票公告列表.csv");
            CsvWriter resultEW = new CsvWriter(resultFilePath, resultColumnDic);
            return resultEW;
        }         

        private void GetGongGaoListAllPagesToCsv(IListSheet listSheet)
        {
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
             
            CsvWriter ew = this.GetCsvWriter();
            Dictionary<string, string> announcementDic = new Dictionary<string, string>();

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"];
                string orgId = row["orgId"];
                string pinyin = row["pinyin"];
                string code = row["code"];
                string zwjc = row["zwjc"];
                string category = row["category"];
                string stockExchange = row["stockExchange"];

                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string localFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir);
                    TextReader tr = null;

                    try
                    {
                        tr = new StreamReader(localFilePath, Encoding.UTF8);
                        string js = tr.ReadToEnd();

                        JObject rootJo = JObject.Parse(js);
                        JArray itemArrayJsons = rootJo.SelectToken("announcements") as JArray;
                        for (int j = 0; j < itemArrayJsons.Count; j++)
                        {
                            JObject itemJson = itemArrayJsons[j] as JObject;
                            string announcementId = itemJson.GetValue("announcementId").ToString().Trim();
                            string announcementTitle = CommonUtil.HtmlDecode(itemJson.GetValue("announcementTitle").ToString().Trim());
                            string announcementTimeStr = itemJson.GetValue("announcementTime").ToString().Trim();
                            string adjunctType = itemJson.GetValue("adjunctType").ToString().Trim();
                            string adjunctUrl = itemJson.GetValue("adjunctUrl").ToString().Trim();

                            DateTime announcementTime = (new DateTime(1970, 1, 1)).AddMilliseconds(long.Parse(announcementTimeStr)).ToLocalTime();
                            string outAnnouncementTimeStr = announcementTime.ToString("yyyy-MM-dd HH:mm:ss");

                            if (!announcementDic.ContainsKey(announcementId))
                            {
                                announcementDic.Add(announcementId, null);

                                Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                f2vs.Add("fileUrl", "http://www.cninfo.com.cn/" + adjunctUrl);
                                f2vs.Add("announcementId", announcementId);
                                f2vs.Add("zwjc", zwjc);
                                f2vs.Add("code", code);
                                f2vs.Add("pinyin", pinyin);
                                f2vs.Add("orgId", orgId);
                                f2vs.Add("category", category);
                                f2vs.Add("stockExchange", stockExchange);
                                f2vs.Add("announcementTitle", announcementTitle);
                                f2vs.Add("announcementTime", outAnnouncementTimeStr);
                                f2vs.Add("adjunctType", adjunctType);
                                ew.AddRow(f2vs);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                    finally
                    {
                        if (tr != null)
                        {
                            tr.Close();
                            tr.Dispose();
                        }
                    }
                }
            }
            ew.SaveToDisk();
        }
         
    }
}