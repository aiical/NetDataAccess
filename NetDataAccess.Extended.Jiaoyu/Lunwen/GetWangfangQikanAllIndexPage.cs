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
    public class GetWangfangQikanAllIndexPage : ExternalRunWebPage
    { 
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            try
            { 
                this.GetAllPerioIndexPageUrls(listSheet);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return true;
        }

        #region 期刊每期所有目录页
        private CsvWriter GetAllPerioIndexPageCsvWriter(int allListFileIndex)
        {
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[] {
                "id",
                "publish_year",
                "fund_info02",
                "page_range",
                "keywords",
                "auto_keys",
                "page_cnt",
                "doc_num",
                "perio_id",
                "language",
                "refdoc_cnt",
                "abstract_url",
                "scholar_id",
                "auto_classcode",
                "authors_name",
                "share_num",
                "trans_column",
                "is_fulltext",
                "issue_num",
                "pro_pub_date",
                "hxkbj_pku",
                "perio_title02",
                "cite_num",
                "unit_name",
                "linkdoc_cnt",
                "issn",
                "unit_name02",
                "data_state",
                "random_id",
                "cited_cnt",
                "doi",
                "fund_info",
                "trans_authors",
                "literature_code",
                "data_sort",
                "new_org",
                "core_perio",
                "publish_year02",
                "auth_area",
                "article_id",
                "tag_num",
                "abstract_reading_num",
                "auto_classcode_level",
                "first_authors",
                "full_pubdate",
                "hxkbj_istic",
                "common_year",
                "authors_unit",
                "thirdparty_links_num",
                "abst_webdate",
                "article_seq",
                "import_num",
                "common_sort_time",
                "issue_id",
                "full_url",
                "orig_pub_date",
                "source_db",
                "column_name",
                "cn",
                "collection_num",
                "download_num",
                "orig_classcode",
                "service_model",
                "first_publish",
                "is_oa",
                "subject_class_codes",
                "fulltext_reading_num",
                "note_num",
                "updatetime",
                "head_words",
                "subject_classcode_level",
                "trans_title",
                "perio_title_en",
                "title",
                "summary",
                "perio_title",
                "class_type",
                "doct_collect"

            });
            string resultFilePath = Path.Combine(exportDir, "万方期刊_期刊每期目录_" + allListFileIndex.ToString() + ".csv");
            CsvWriter resultEW = new CsvWriter(resultFilePath, resultColumnDic);
            return resultEW;
        }

        /// <summary>
        /// 期刊每期目录首页
        /// </summary>
        /// <param name="listSheet"></param>
        private void GetAllPerioIndexPageUrls(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            int allListFileIndex = 1;
            CsvWriter ew = null;
            Dictionary<string, string> idDic = new Dictionary<string, string>(); 
            int paperCount = 0;
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                if (paperCount % 1000 == 0)
                {
                    this.RunPage.InvokeAppendLogText("已处理到: fileIndex = " + allListFileIndex.ToString() + ", paperIndex = " + paperCount.ToString(), LogLevelType.System, true);
                }

                if (paperCount >= 500000)
                {
                    if (ew != null)
                    {
                        ew.SaveToDisk();
                    }
                    ew = this.GetAllPerioIndexPageCsvWriter(allListFileIndex);
                    allListFileIndex++;
                    paperCount = 0;
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
                        if (itemJsonArray != null && itemJsonArray.Count > 0)
                        {
                            for (int j = 0; j < itemJsonArray.Count; j++)
                            {
                                JObject itemJson = itemJsonArray[j] as JObject;
                                string id = itemJson.GetValue("id").ToString();
                                if (!idDic.ContainsKey(id))
                                {
                                    idDic.Add(id, null);

                                    Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                    this.GetAttributeValue(itemJson, "id", f2vs);
                                    this.GetAttributeValue(itemJson, "publish_year", f2vs);
                                    this.GetAttributeValue(itemJson, "fund_info02", f2vs);
                                    this.GetAttributeValue(itemJson, "page_range", f2vs);
                                    this.GetAttributeValue(itemJson, "keywords", f2vs);
                                    this.GetAttributeValue(itemJson, "auto_keys", f2vs);
                                    this.GetAttributeValue(itemJson, "page_cnt", f2vs);
                                    this.GetAttributeValue(itemJson, "doc_num", f2vs);
                                    this.GetAttributeValue(itemJson, "perio_id", f2vs);
                                    this.GetAttributeValue(itemJson, "language", f2vs);
                                    this.GetAttributeValue(itemJson, "refdoc_cnt", f2vs);
                                    this.GetAttributeValue(itemJson, "abstract_url", f2vs);
                                    this.GetAttributeValue(itemJson, "scholar_id", f2vs);
                                    this.GetAttributeValue(itemJson, "auto_classcode", f2vs);
                                    this.GetAttributeValue(itemJson, "authors_name", f2vs);
                                    this.GetAttributeValue(itemJson, "share_num", f2vs);
                                    this.GetAttributeValue(itemJson, "trans_column", f2vs);
                                    this.GetAttributeValue(itemJson, "is_fulltext", f2vs);
                                    this.GetAttributeValue(itemJson, "issue_num", f2vs);
                                    this.GetAttributeValue(itemJson, "pro_pub_date", f2vs);
                                    this.GetAttributeValue(itemJson, "hxkbj_pku", f2vs);
                                    this.GetAttributeValue(itemJson, "perio_title02", f2vs);
                                    this.GetAttributeValue(itemJson, "cite_num", f2vs);
                                    this.GetAttributeValue(itemJson, "unit_name", f2vs);
                                    this.GetAttributeValue(itemJson, "linkdoc_cnt", f2vs);
                                    this.GetAttributeValue(itemJson, "issn", f2vs);
                                    this.GetAttributeValue(itemJson, "unit_name02", f2vs);
                                    this.GetAttributeValue(itemJson, "data_state", f2vs);
                                    this.GetAttributeValue(itemJson, "random_id", f2vs);
                                    this.GetAttributeValue(itemJson, "cited_cnt", f2vs);
                                    this.GetAttributeValue(itemJson, "doi", f2vs);
                                    this.GetAttributeValue(itemJson, "fund_info", f2vs);
                                    this.GetAttributeValue(itemJson, "trans_authors", f2vs);
                                    this.GetAttributeValue(itemJson, "literature_code", f2vs);
                                    this.GetAttributeValue(itemJson, "data_sort", f2vs);
                                    this.GetAttributeValue(itemJson, "new_org", f2vs);
                                    this.GetAttributeValue(itemJson, "core_perio", f2vs);
                                    this.GetAttributeValue(itemJson, "publish_year02", f2vs);
                                    this.GetAttributeValue(itemJson, "auth_area", f2vs);
                                    this.GetAttributeValue(itemJson, "article_id", f2vs);
                                    this.GetAttributeValue(itemJson, "tag_num", f2vs);
                                    this.GetAttributeValue(itemJson, "abstract_reading_num", f2vs);
                                    this.GetAttributeValue(itemJson, "auto_classcode_level", f2vs);
                                    this.GetAttributeValue(itemJson, "first_authors", f2vs);
                                    this.GetAttributeValue(itemJson, "full_pubdate", f2vs);
                                    this.GetAttributeValue(itemJson, "hxkbj_istic", f2vs);
                                    this.GetAttributeValue(itemJson, "common_year", f2vs);
                                    this.GetAttributeValue(itemJson, "authors_unit", f2vs);
                                    this.GetAttributeValue(itemJson, "thirdparty_links_num", f2vs);
                                    this.GetAttributeValue(itemJson, "abst_webdate", f2vs);
                                    this.GetAttributeValue(itemJson, "article_seq", f2vs);
                                    this.GetAttributeValue(itemJson, "import_num", f2vs);
                                    this.GetAttributeValue(itemJson, "common_sort_time", f2vs);
                                    this.GetAttributeValue(itemJson, "issue_id", f2vs);
                                    this.GetAttributeValue(itemJson, "full_url", f2vs);
                                    this.GetAttributeValue(itemJson, "orig_pub_date", f2vs);
                                    this.GetAttributeValue(itemJson, "source_db", f2vs);
                                    this.GetAttributeValue(itemJson, "column_name", f2vs);
                                    this.GetAttributeValue(itemJson, "cn", f2vs);
                                    this.GetAttributeValue(itemJson, "collection_num", f2vs);
                                    this.GetAttributeValue(itemJson, "download_num", f2vs);
                                    this.GetAttributeValue(itemJson, "orig_classcode", f2vs);
                                    this.GetAttributeValue(itemJson, "service_model", f2vs);
                                    this.GetAttributeValue(itemJson, "first_publish", f2vs);
                                    this.GetAttributeValue(itemJson, "is_oa", f2vs);
                                    this.GetAttributeValue(itemJson, "subject_class_codes", f2vs);
                                    this.GetAttributeValue(itemJson, "fulltext_reading_num", f2vs);
                                    this.GetAttributeValue(itemJson, "note_num", f2vs);
                                    this.GetAttributeValue(itemJson, "updatetime", f2vs);
                                    this.GetAttributeValue(itemJson, "head_words", f2vs);
                                    this.GetAttributeValue(itemJson, "subject_classcode_level", f2vs);
                                    this.GetAttributeValue(itemJson, "trans_title", f2vs);
                                    this.GetAttributeValue(itemJson, "perio_title_en", f2vs);
                                    this.GetAttributeValue(itemJson, "title", f2vs);
                                    this.GetAttributeValue(itemJson, "summary", f2vs);
                                    this.GetAttributeValue(itemJson, "perio_title", f2vs);
                                    this.GetAttributeValue(itemJson, "class_type", f2vs);
                                    this.GetAttributeValue(itemJson, "doct_collect", f2vs);
                                    
                                    paperCount++;
                                    
                                    ew.AddRow(f2vs);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        this.RunPage.InvokeAppendLogText(ex.Message + ". detailUrl = " + detailUrl, LogLevelType.Error, true);
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