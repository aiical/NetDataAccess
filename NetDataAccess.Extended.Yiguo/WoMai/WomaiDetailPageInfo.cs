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
using NetDataAccess.Base.Writer;
using NetDataAccess.Base.Reader;
using NetDataAccess.Base.CsvHelper;
using NetDataAccess.Base.DB;
using HtmlAgilityPack;

namespace NetDataAccess.Extended.Yiguo
{
    public class WomaiDetailPageInfo : CustomProgramBase
    {
        public bool Run(string parameters, IListSheet listSheet)
        {
            return this.GeneratePageInfo(listSheet);
        }

        private bool GeneratePageInfo(IListSheet listSheet)
        {
            bool succeed = true;
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            string exportDir = this.RunPage.GetExportDir();
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
            string resultFilePath = Path.Combine(exportDir, "我买网获取所有详情页_Redirect.xlsx");

            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic);

            GetPageUrl(listSheet, pageSourceDir, resultEW);

            resultEW.SaveToDisk();

            return succeed;
        }

        private void GetPageUrl(IListSheet listSheet, string pageSourceDir, ExcelWriter resultEW)
        {
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string pageUrl = listSheet.PageUrlList[i];
                string pageName = listSheet.PageNameList[i];
                string localFilePath = this.RunPage.GetFilePath(pageUrl, pageSourceDir);  

                TextReader tr = null;

                try
                {
                    tr = new StreamReader(localFilePath, Encoding.GetEncoding("GBK"));
                    string webPageHtml = tr.ReadToEnd().Trim();
                    string detailPageUrl = pageUrl;
                    if (webPageHtml.StartsWith("<script>location = \""))
                    {
                        string url  =  webPageHtml.Replace("\";</script>", "").Replace("<script>location = \"", "");
                        detailPageUrl = url.StartsWith("http") ? url : ("http://www.womai.com" + url);
                    } 

                    Dictionary<string, string> f2vs = new Dictionary<string, string>();
                    f2vs.Add("detailPageUrl", detailPageUrl);
                    f2vs.Add("detailPageName", pageName);
                    f2vs.Add("cookie", row["cookie"]);
                    f2vs.Add("category1Code", row["category1Code"]);
                    f2vs.Add("category2Code", row["category2Code"]);
                    f2vs.Add("category1Name", row["category1Name"]);
                    f2vs.Add("category2Name", row["category2Name"]);
                    f2vs.Add("district", row["district"]);
                    f2vs.Add("productSysNo", row["productSysNo"]);
                    resultEW.AddRow(f2vs);

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
         
    }
}