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

namespace NetDataAccess.Extended.GuPiao
{
    public class GetHuShenNianbaoList : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            this.GetGuPiaoNianbaoPageUrls(listSheet);
            return true;
        }
        
        private ExcelWriter GetDetailPageExcelWriter(int fileIndex)
        {
            String exportDir = this.Parameters.Trim();

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
            string resultFilePath = Path.Combine(exportDir, "沪深股票年报内容页_" + fileIndex.ToString() + ".xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }

        private void GetGuPiaoNianbaoPageUrls(IListSheet listSheet)
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
                string detailPageName = row["detailPageName"];
                string pinyin = row["pinyin"];
                string zwjc = row["zwjc"];
                string code = row["code"];
                string orgId = row["orgId"];
                string stockExchange = row["stockExchange"];
                string category = row["category"];
                string announcementTitle = row["announcementTitle"].Trim();
                string announcementTime = row["announcementTime"];
                string adjunctType = row["adjunctType"]; 

                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {

                    try
                    {
                        if ((announcementTitle == "2017年年度报告") || (announcementTitle.StartsWith("2017年年度报告（") && !announcementTitle.Contains("英文") && !announcementTitle.Contains("摘要")))
                        {
                            Dictionary<string, string> f2vs = new Dictionary<string, string>();
                            f2vs.Add("detailPageUrl", detailUrl);
                            f2vs.Add("detailPageName", detailPageName);
                            f2vs.Add("zwjc", zwjc);
                            f2vs.Add("code", code);
                            f2vs.Add("pinyin", pinyin);
                            f2vs.Add("orgId", orgId);
                            f2vs.Add("category", category);
                            f2vs.Add("stockExchange", stockExchange);
                            f2vs.Add("announcementTitle", announcementTitle);
                            f2vs.Add("announcementTime", announcementTime);
                            f2vs.Add("adjunctType", adjunctType);
                            ew.AddRow(f2vs);
                        }
                        else if ((announcementTitle == "2016年年度报告") || (announcementTitle.StartsWith("2016年年度报告（") && !announcementTitle.Contains("英文") && !announcementTitle.Contains("摘要")))
                        {
                            Dictionary<string, string> f2vs = new Dictionary<string, string>();
                            f2vs.Add("detailPageUrl", detailUrl);
                            f2vs.Add("detailPageName", detailPageName);
                            f2vs.Add("zwjc", zwjc);
                            f2vs.Add("code", code);
                            f2vs.Add("pinyin", pinyin);
                            f2vs.Add("orgId", orgId);
                            f2vs.Add("category", category);
                            f2vs.Add("stockExchange", stockExchange);
                            f2vs.Add("announcementTitle", announcementTitle);
                            f2vs.Add("announcementTime", announcementTime);
                            f2vs.Add("adjunctType", adjunctType);
                            ew.AddRow(f2vs);
                        }
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                    finally
                    { 
                    }
                }
            }
            ew.SaveToDisk();
        }
    }
}