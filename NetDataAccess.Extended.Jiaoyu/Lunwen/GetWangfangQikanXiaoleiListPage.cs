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
    public class GetWangfangQikanXiaoleiListPage : ExternalRunWebPage
    { 
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            try
            {
                this.GetPeriodicalCategoryList(listSheet);
                this.GetPeriodicalListPageUrls(listSheet); 
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return true;
        }

        private void GetPeriodicalCategoryList(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>(); 
            resultColumnDic.Add("cate1", 1);
            resultColumnDic.Add("cateId1", 2);
            resultColumnDic.Add("cate2", 3);
            resultColumnDic.Add("cateId2", 4);
            resultColumnDic.Add("periodicalCount", 5);
            string resultFilePath = Path.Combine(exportDir, "万方期刊_期刊分类.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"];

                string cate1 = row["cate1"];
                string cateId1 = row["cateId1"];
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string localFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir);

                    try
                    {
                        string pageFileText = FileHelper.GetTextFromFile(localFilePath);
                        JArray itemJsonArray = JArray.Parse(pageFileText);


                        for (int j = 0; j < itemJsonArray.Count; j++)
                        {
                            JObject itemJson = itemJsonArray[j] as JObject;
                            string cateId2 = itemJson.GetValue("id").ToString();
                            string cate2 = itemJson.GetValue("showName").ToString().Trim();
                            int periodicalCount = int.Parse(itemJson.GetValue("count").ToString().Trim());
                            Dictionary<string, string> f2vs = new Dictionary<string, string>();
                            f2vs.Add("cate1", cate1);
                            f2vs.Add("cateId1", cateId1);
                            f2vs.Add("cate2", cate2);
                            f2vs.Add("cateId2", cateId2);
                            f2vs.Add("periodicalCount", periodicalCount.ToString());
                            resultEW.AddRow(f2vs);
                        }

                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
            }
            resultEW.SaveToDisk();
        }

        private void GetPeriodicalListPageUrls(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("cate1", 5);
            resultColumnDic.Add("cateId1", 6);
            resultColumnDic.Add("cate2", 7);
            resultColumnDic.Add("cateId2", 8);
            resultColumnDic.Add("pageIndex", 9);
            string resultFilePath = Path.Combine(exportDir, "万方期刊_期刊列表.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"];

                string cate1 = row["cate1"];
                string cateId1 = row["cateId1"]; 
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string localFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir);

                    try
                    {
                        string pageFileText = FileHelper.GetTextFromFile(localFilePath);
                        JArray itemJsonArray = JArray.Parse(pageFileText);


                        for (int j = 0; j < itemJsonArray.Count; j++)
                        {
                            JObject itemJson = itemJsonArray[j] as JObject;
                            string cateId2 = itemJson.GetValue("id").ToString();
                            string cate2 = itemJson.GetValue("showName").ToString().Trim();
                            int periodicalCount = int.Parse(itemJson.GetValue("count").ToString().Trim());  
                            int pageCount = periodicalCount == 0 ? 0 : (periodicalCount / 20 + 1);
                            for (int k = 0; k < pageCount; k++)
                            {
                                string newUrl = "http://www.wanfangdata.com.cn/perio/page.do?page=" + (k + 1).ToString() + "&pageSize=20&selectOrder=affectoi&fmList=" + cateId2 + "&a_title=&core=&fromData=WF&included=&publishyear=&isfirst=";
                                Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                f2vs.Add("detailPageUrl", newUrl);
                                f2vs.Add("detailPageName", newUrl);
                                f2vs.Add("cate1", cate1);
                                f2vs.Add("cateId1", cateId1);
                                f2vs.Add("cate2", cate2);
                                f2vs.Add("cateId2", cateId2);
                                f2vs.Add("pageIndex", (k + 1).ToString());
                                resultEW.AddRow(f2vs);
                            }
                        }

                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
            }
            resultEW.SaveToDisk();
        } 
    }
}