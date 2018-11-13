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
    public class GetWangfangQikanFirstIndexPage : ExternalRunWebPage
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

        #region 期刊每期目录首页
        private ExcelWriter GetAllPerioIndexPageExcelWriter(int allListFileIndex)
        {
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[] { "detailPageUrl", "detailPageName","cookie", "grabStatus",  "giveUpGrab", "perio_id", "issue_num", "publish_year", "perio_title", "pageIndex" });
            string resultFilePath = Path.Combine(exportDir, "万方期刊_期刊每期所有目录页_" + allListFileIndex.ToString() + ".xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
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
                    ew = this.GetAllPerioIndexPageExcelWriter(allListFileIndex);
                    allListFileIndex++;
                }

                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"]; 
                string perio_id = row["perio_id"];
                string issue_num = row["issue_num"];
                string publish_year = row["publish_year"];
                string perio_title = row["perio_title"]; 

                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string localFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir);

                    try
                    {
                        string pageFileText = FileHelper.GetTextFromFile(localFilePath);
                        int pageTotal = int.Parse(JObject.Parse(pageFileText).GetValue("pageTotal").ToString());
                        for (int j = 1; j <= pageTotal; j++)
                        {
                            string indexPageUrl = "http://www.wanfangdata.com.cn/perio/articleList.do?page=" + (j + 1).ToString() + "&pageSize=10&issue_num=" + issue_num + "&publish_year=" + publish_year + "&article_start=&title_article=&perio_id=" + perio_id;
                            if (!urlDic.ContainsKey(indexPageUrl))
                            {
                                urlDic.Add(indexPageUrl, null);

                                Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                f2vs.Add("detailPageUrl", indexPageUrl);
                                f2vs.Add("detailPageName", indexPageUrl);
                                f2vs.Add("perio_id", perio_id);
                                f2vs.Add("issue_num", issue_num);
                                f2vs.Add("publish_year", publish_year);
                                f2vs.Add("perio_title", perio_title);
                                f2vs.Add("pageIndex", j.ToString());
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