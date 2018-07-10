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

namespace NetDataAccess.Extended.Celloud
{
    public class OalibGetListPage : CustomProgramBase
    {
        public bool Run(string parameters, IListSheet listSheet)
        {
            return GetAllDetailPageUrl(listSheet);
        }
        private bool GetAllDetailPageUrl(IListSheet listSheet)
        {
            string exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                "detailPageUrl",
                "detailPageName", 
                "cookie",
                "grabStatus", 
                "giveUpGrab",
                "paperName",
                "species", 
                "year"});
            string resultFilePath = Path.Combine(exportDir, "oalib获取摘要页面.xlsx");
            
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic);

            Dictionary<string, string> goodsDic = new Dictionary<string, string>();

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;
            string categoryNameColumnName = SysConfig.DetailPageNameFieldName;

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string url = row[detailPageUrlColumnName];
                    string cookie = row["cookie"];
                    string species = row["species"].Trim(); 
                    string year = row["year"].Trim();
                    string localFilePath = this.RunPage.GetFilePath(url, pageSourceDir); 

                    try
                    {
                        {
                            HtmlNode.ElementsFlags.Remove("form");
                            HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                            HtmlNodeCollection allItemParentNodes = htmlDoc.DocumentNode.SelectNodes("//html/body/form[1]/div/center/table/tbody/tr[2]/td/table/tr/td[2]/div[3]/table/tr");
                            if (allItemParentNodes != null)
                            {
                                foreach (HtmlNode itemParentNode in allItemParentNodes)
                                {
                                    HtmlNode itemNode = itemParentNode.SelectSingleNode("./td/table/tr/td/span[1]/a[1]");
                                    string detailPageUrl = itemNode.Attributes["href"].Value;
                                    string detailPageName = detailPageUrl.Substring(detailPageUrl.LastIndexOf("/") + 1) + "_" + species;
                                    string paperName = CommonUtil.HtmlDecode(itemNode.InnerText.Trim());

                                    if (!goodsDic.ContainsKey(detailPageName))
                                    {
                                        goodsDic.Add(detailPageName, null);
                                        Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                        f2vs.Add("detailPageUrl", detailPageUrl);
                                        f2vs.Add("detailPageName", detailPageName);
                                        f2vs.Add("cookie", cookie);
                                        f2vs.Add("paperName", paperName);
                                        f2vs.Add("species", species); 
                                        f2vs.Add("year", year);
                                        resultEW.AddRow(f2vs);
                                    }
                                    else
                                    { 
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        this.RunPage.InvokeAppendLogText("读取出错.  " + ex.Message + " LocalPath = " + localFilePath, LogLevelType.Error, true);
                        throw ex;
                    }
                }
            }
            resultEW.SaveToDisk(); 
            return true;
        }
    }
}