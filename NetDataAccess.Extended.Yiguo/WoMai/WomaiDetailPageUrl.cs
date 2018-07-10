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

namespace NetDataAccess.Extended.Yiguo
{
    public class WomaiDetailPageUrl : CustomProgramBase
    {
        public bool Run(string parameters, IListSheet listSheet)
        {
            return GetAllDetailPageUrl(listSheet);
        }
        private bool GetAllDetailPageUrl(IListSheet listSheet)
        {
            string exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("productSysNo", 5);
            resultColumnDic.Add("category1Code", 6);
            resultColumnDic.Add("category2Code", 7); 
            resultColumnDic.Add("category1Name", 8);
            resultColumnDic.Add("category2Name", 9); 
            resultColumnDic.Add("district", 10);
            string resultFilePath = Path.Combine(exportDir, "我买网获取所有详情页.xlsx");
            
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
                    string category1Code = row["category1Code"];
                    string category2Code = row["category2Code"]; 
                    string category1Name = row["category1Name"];
                    string category2Name = row["category2Name"]; 
                    string district = row["district"];
                    string cookie = row["cookie"]; 
                    string localFilePath = this.RunPage.GetFilePath(url, pageSourceDir);
                    TextReader tr = null;

                    try
                    {
                        tr = new StreamReader(localFilePath, Encoding.GetEncoding("GBK"));
                        string webPageHtml = tr.ReadToEnd();

                        HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                        htmlDoc.LoadHtml(webPageHtml);

                        HtmlNodeCollection allDetailPageNodes = htmlDoc.DocumentNode.SelectNodes("//div[@class=\"product-list\"]/ul/li/div[@class=\"tab-content\"]/div[@class=\"list-title\"]/p/a");

                        foreach (HtmlNode detailPageNode in allDetailPageNodes)
                        {
                            string urlStr = detailPageNode.Attributes["href"].Value;
                            string[] urlSplits = urlStr.Split(new string[] { "-", "." }, StringSplitOptions.RemoveEmptyEntries);
                            string productSysNo = urlSplits[urlSplits.Length - 2];
                            string detailPageName = productSysNo;
                            string detailPageUrl = "http://www.womai.com" + urlStr;

                            if (!goodsDic.ContainsKey(detailPageName))
                            {
                                goodsDic.Add(detailPageName, null);
                                Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                f2vs.Add("detailPageUrl", detailPageUrl);
                                f2vs.Add("detailPageName", detailPageName);
                                f2vs.Add("cookie", cookie);  
                                f2vs.Add("category1Code", category1Code);
                                f2vs.Add("category2Code", category2Code); 
                                f2vs.Add("category1Name", category1Name);
                                f2vs.Add("category2Name", category2Name); 
                                f2vs.Add("district", district);
                                f2vs.Add("productSysNo", productSysNo);
                                resultEW.AddRow(f2vs);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        if (tr != null)
                        {
                            tr.Dispose();
                            tr = null;
                        }
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