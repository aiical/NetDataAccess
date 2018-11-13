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

namespace NetDataAccess.Extended.Jiaoyu.Lunwen
{
    public class GetWangfangQikanListPage : ExternalRunWebPage
    { 
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            try
            {
                //this.GetPeriodicalInfo(listSheet);
                //this.GetPeriodicalPerioIssueInfo(listSheet);
                this.GetAllPerioFirstIndexPageUrls(listSheet);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return true;
        }

        #region 每个期刊的信息
        private void GetPeriodicalInfo(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[] { "id", "core_perio", "avg_perio_down", "start_year02", "start_year", "issue_postcode", "perio_format", "fax", "perio_id", "language", "tag_num", "major_editor", "abstract_reading_num", "thirdparty_links_num", "import_num", "email", "share_num", "classcode_level", "publish_cycle", "address", "pinyin_title", "avg_article_down", "hostunit_name", "hostunit_area", "director", "main_column", "telephone", "country_code", "affectoi", "issn", "cn", "source_db", "dep_name", "postcode", "collection_num", "win_prize", "cite_num", "perio_title02", "download_num", "first_publish", "data_state", "article_num", "ef_name", "release_cycle", "fulltext_reading_num", "note_num", "end_year", "class_code", "end_issue", "trans_title", "perio_desc", "perio_title", "keywords", "summary", "cate1", "cateId1", "cate2", "cateId2" });
            string resultFilePath = Path.Combine(exportDir, "万方期刊_期刊信息详情.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"];

                try
                {
                    string cate1 = row["cate1"];
                    string cateId1 = row["cateId1"];
                    string cate2 = row["cate2"];
                    string cateId2 = row["cateId2"];
                    bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                    if (!giveUp)
                    {
                        string localFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir);

                        try
                        {
                            string pageFileText = FileHelper.GetTextFromFile(localFilePath);
                            JArray itemJsonArray = JObject.Parse(pageFileText).GetValue("pageRow") as JArray;


                            for (int j = 0; j < itemJsonArray.Count; j++)
                            {
                                Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                JObject itemJson = itemJsonArray[j] as JObject;
                                f2vs.Add("cate1", cate1);
                                f2vs.Add("cateId1", cateId1);
                                f2vs.Add("cate2", cate2);
                                f2vs.Add("cateId2", cateId2);

                                this.GetAttributeValue(itemJson, "id", f2vs);
                                this.GetAttributeValue(itemJson, "core_perio", f2vs);
                                this.GetAttributeValue(itemJson, "avg_perio_down", f2vs);
                                this.GetAttributeValue(itemJson, "start_year02", f2vs);
                                this.GetAttributeValue(itemJson, "start_year", f2vs);
                                this.GetAttributeValue(itemJson, "issue_postcode", f2vs);
                                this.GetAttributeValue(itemJson, "perio_format", f2vs);
                                this.GetAttributeValue(itemJson, "fax", f2vs);
                                this.GetAttributeValue(itemJson, "perio_id", f2vs);
                                this.GetAttributeValue(itemJson, "language", f2vs);
                                this.GetAttributeValue(itemJson, "tag_num", f2vs);
                                this.GetAttributeValue(itemJson, "major_editor", f2vs);
                                this.GetAttributeValue(itemJson, "abstract_reading_num", f2vs);
                                this.GetAttributeValue(itemJson, "thirdparty_links_num", f2vs);
                                this.GetAttributeValue(itemJson, "import_num", f2vs);
                                this.GetAttributeValue(itemJson, "email", f2vs);
                                this.GetAttributeValue(itemJson, "share_num", f2vs);
                                this.GetAttributeValue(itemJson, "classcode_level", f2vs);
                                this.GetAttributeValue(itemJson, "publish_cycle", f2vs);
                                this.GetAttributeValue(itemJson, "address", f2vs);
                                this.GetAttributeValue(itemJson, "pinyin_title", f2vs);
                                this.GetAttributeValue(itemJson, "avg_article_down", f2vs);
                                this.GetAttributeValue(itemJson, "hostunit_name", f2vs);
                                this.GetAttributeValue(itemJson, "hostunit_area", f2vs);
                                this.GetAttributeValue(itemJson, "director", f2vs);
                                this.GetAttributeValue(itemJson, "main_column", f2vs);
                                this.GetAttributeValue(itemJson, "telephone", f2vs);
                                this.GetAttributeValue(itemJson, "country_code", f2vs);
                                this.GetAttributeValue(itemJson, "affectoi", f2vs);
                                this.GetAttributeValue(itemJson, "issn", f2vs);
                                this.GetAttributeValue(itemJson, "cn", f2vs);
                                this.GetAttributeValue(itemJson, "source_db", f2vs);
                                this.GetAttributeValue(itemJson, "dep_name", f2vs);
                                this.GetAttributeValue(itemJson, "postcode", f2vs);
                                this.GetAttributeValue(itemJson, "collection_num", f2vs);
                                this.GetAttributeValue(itemJson, "win_prize", f2vs);
                                this.GetAttributeValue(itemJson, "cite_num", f2vs);
                                this.GetAttributeValue(itemJson, "perio_title02", f2vs);
                                this.GetAttributeValue(itemJson, "download_num", f2vs);
                                this.GetAttributeValue(itemJson, "first_publish", f2vs);
                                this.GetAttributeValue(itemJson, "data_state", f2vs);
                                this.GetAttributeValue(itemJson, "article_num", f2vs);
                                this.GetAttributeValue(itemJson, "ef_name", f2vs);
                                this.GetAttributeValue(itemJson, "release_cycle", f2vs);
                                this.GetAttributeValue(itemJson, "fulltext_reading_num", f2vs);
                                this.GetAttributeValue(itemJson, "note_num", f2vs);
                                this.GetAttributeValue(itemJson, "end_year", f2vs);
                                this.GetAttributeValue(itemJson, "class_code", f2vs);
                                this.GetAttributeValue(itemJson, "end_issue", f2vs);
                                this.GetAttributeValue(itemJson, "trans_title", f2vs);
                                this.GetAttributeValue(itemJson, "perio_desc", f2vs);
                                this.GetAttributeValue(itemJson, "perio_title", f2vs);
                                this.GetAttributeValue(itemJson, "keywords", f2vs);
                                this.GetAttributeValue(itemJson, "summary", f2vs);

                                resultEW.AddRow(f2vs);
                            }

                        }
                        catch (Exception ex)
                        { 
                            throw ex;
                        }
                    }
                }
                catch (Exception ex)
                {
                    this.RunPage.InvokeAppendLogText(ex.Message + ". detailUrl" + detailUrl, LogLevelType.Error, true);
                    throw ex;
                }
            }
            resultEW.SaveToDisk();
        }
        #endregion

        #region 期刊每期的信息
        private ExcelWriter GetAllPerioIssueInfoExcelWriter(int allListFileIndex)
        {
            String exportDir = this.RunPage.GetExportDir();
            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[] { "perioId", "publish_year", "trans_title", "issue_id", "show_issue_num", "page_cnt", "issue_num", "perio_id", "orig_catalog", "volume", "catalog_url", "total_issue", "special_title", "issue_cover", "id", "perio_title" });
            string resultFilePath = Path.Combine(exportDir, "万方期刊_期刊每期列表_" + allListFileIndex.ToString() + ".xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }

        /// <summary>
        /// 获取每一期的基本信息
        /// </summary>
        /// <param name="listSheet"></param>
        private void GetPeriodicalPerioIssueInfo(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            int allListFileIndex = 1;
            ExcelWriter ew = null;
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                if (ew == null || ew.RowCount > 500000)
                {
                    if (ew != null)
                    {
                        ew.SaveToDisk();
                    }
                    ew = this.GetAllPerioIssueInfoExcelWriter(allListFileIndex);
                    allListFileIndex++;
                }

                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"];

                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string localFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir);

                    try
                    {
                        string pageFileText = FileHelper.GetTextFromFile(localFilePath);
                        JArray itemJsonArray = JObject.Parse(pageFileText).GetValue("pageRow") as JArray;

                        if (itemJsonArray != null)
                        {
                            for (int j = 0; j < itemJsonArray.Count; j++)
                            {
                                JObject itemJson = itemJsonArray[j] as JObject;
                                string perioId = itemJson.GetValue("id").ToString().Trim();
                                JObject opJson = itemJson.GetValue("op") as JObject;
                                try
                                {
                                    if (opJson != null)
                                    {
                                        JArray opItemsArray = opJson.GetValue("perioIssue") as JArray;
                                        //每一期
                                        if (opItemsArray != null)
                                        {
                                            for (int k = 0; k < opItemsArray.Count; k++)
                                            {
                                                JObject opItemJson = opItemsArray[k] as JObject;

                                                Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                                f2vs.Add("perioId", perioId);
                                                this.GetAttributeValue(opItemJson, "publish_year", f2vs);
                                                this.GetAttributeValue(opItemJson, "trans_title", f2vs);
                                                this.GetAttributeValue(opItemJson, "issue_id", f2vs);
                                                this.GetAttributeValue(opItemJson, "show_issue_num", f2vs);
                                                this.GetAttributeValue(opItemJson, "page_cnt", f2vs);
                                                this.GetAttributeValue(opItemJson, "issue_num", f2vs);
                                                this.GetAttributeValue(opItemJson, "perio_id", f2vs);
                                                this.GetAttributeValue(opItemJson, "orig_catalog", f2vs);
                                                this.GetAttributeValue(opItemJson, "volume", f2vs);
                                                this.GetAttributeValue(opItemJson, "catalog_url", f2vs);
                                                this.GetAttributeValue(opItemJson, "total_issue", f2vs);
                                                this.GetAttributeValue(opItemJson, "special_title", f2vs);
                                                this.GetAttributeValue(opItemJson, "issue_cover", f2vs);
                                                this.GetAttributeValue(opItemJson, "id", f2vs);
                                                this.GetAttributeValue(opItemJson, "perio_title", f2vs);

                                                ew.AddRow(f2vs);
                                            }
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    throw ex;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
            }
            ew.SaveToDisk();
        }
        #endregion

        #region 期刊每期目录首页
        private ExcelWriter GetAllPerioFirstIndexPageExcelWriter(int allListFileIndex)
        {
            String exportDir = this.RunPage.GetExportDir();
            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[] { "detailPageUrl", "detailPageName","cookie", "grabStatus",  "giveUpGrab", "perio_id", "issue_num", "publish_year", "perio_title", "pageIndex" });
            string resultFilePath = Path.Combine(exportDir, "万方期刊_期刊每期目录首页_" + allListFileIndex.ToString() + ".xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }

        /// <summary>
        /// 期刊每期目录首页
        /// </summary>
        /// <param name="listSheet"></param>
        private void GetAllPerioFirstIndexPageUrls(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            int allListFileIndex = 1;
            ExcelWriter ew = null;
            Dictionary<string, string> urlDic = new Dictionary<string, string>();
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                if (ew == null || ew.RowCount > 500000)
                {
                    if (ew != null)
                    {
                        ew.SaveToDisk();
                    }
                    ew = this.GetAllPerioFirstIndexPageExcelWriter(allListFileIndex);
                    allListFileIndex++;
                }

                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"];

                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string localFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir);

                    try
                    {
                        string pageFileText = FileHelper.GetTextFromFile(localFilePath);
                        JArray itemJsonArray = JObject.Parse(pageFileText).GetValue("pageRow") as JArray;

                        if (itemJsonArray != null)
                        {
                            for (int j = 0; j < itemJsonArray.Count; j++)
                            {
                                JObject itemJson = itemJsonArray[j] as JObject;
                                string perioId = itemJson.GetValue("id").ToString().Trim();
                                JObject opJson = itemJson.GetValue("op") as JObject;
                                try
                                {
                                    if (opJson != null)
                                    {
                                        JArray opItemsArray = opJson.GetValue("perioIssue") as JArray;
                                        //每一期
                                        if (opItemsArray != null)
                                        {
                                            for (int k = 0; k < opItemsArray.Count; k++)
                                            {
                                                JObject opItemJson = opItemsArray[k] as JObject;
                                                try
                                                {
                                                    string issue_num = this.GetAttributeValue(opItemJson, "issue_num");
                                                    string publish_year = this.GetAttributeValue(opItemJson, "publish_year");
                                                    string perio_id = this.GetAttributeValue(opItemJson, "perio_id");
                                                    string perio_title = this.GetAttributeValue(opItemJson, "perio_title");
                                                    if (issue_num != null && publish_year != null && perio_id != null && perio_title != null)
                                                    {
                                                        string firstIndexPageUrl = "http://www.wanfangdata.com.cn/perio/articleList.do?page=1&pageSize=10&issue_num=" + issue_num + "&publish_year=" + publish_year + "&article_start=&title_article=&perio_id=" + perio_id;
                                                        if (!urlDic.ContainsKey(firstIndexPageUrl))
                                                        {
                                                            urlDic.Add(firstIndexPageUrl, null);

                                                            Dictionary<string, string> f2vs = new Dictionary<string, string>();

                                                            f2vs.Add("detailPageUrl", firstIndexPageUrl);
                                                            f2vs.Add("detailPageName", firstIndexPageUrl);
                                                            f2vs.Add("perio_id", perio_id);
                                                            f2vs.Add("issue_num", issue_num);
                                                            f2vs.Add("publish_year", publish_year);
                                                            f2vs.Add("perio_title", perio_title);
                                                            f2vs.Add("pageIndex", "1");
                                                            ew.AddRow(f2vs);
                                                        }
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
                                catch (Exception ex)
                                {
                                    throw ex;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
            }
            ew.SaveToDisk();
        }
        #endregion

        private void GetAttributeValue(JObject itemJson, string propertyName, Dictionary<string, string> row)
        {
            JToken jt = null;
            if (itemJson.TryGetValue(propertyName, out jt))
            {
                row.Add(propertyName, (jt == null ? "" : jt.ToString()));
            }
        }

        private string GetAttributeValue(JObject itemJson, string propertyName)
        {
            JToken jt = null;
            if (itemJson.TryGetValue(propertyName, out jt))
            {
                return jt == null ? "" : jt.ToString();
            }
            else
            {
                return null;
            }
        }
    }
}