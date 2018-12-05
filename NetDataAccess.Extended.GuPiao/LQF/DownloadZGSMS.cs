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
using NetDataAccess.Base.Reader;

namespace NetDataAccess.Extended.GuPiao.LQF
{
    /// <summary>
    /// 下载招股说明书
    /// </summary>
    public class DownloadZGSMS : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            try
            {
                this.ProcessFiles(listSheet);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return true;
        }

        private void ProcessFiles(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("url", 0);
            resultColumnDic.Add("fileName", 1);
            resultColumnDic.Add("简写", 2);
            resultColumnDic.Add("名称", 3);
            resultColumnDic.Add("编码", 4);
            resultColumnDic.Add("标题", 5);
            string resultFilePath = Path.Combine(exportDir, "招股说明书文件列表.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "Matched", resultColumnDic);

            string exportGGFileDir = Path.Combine(exportDir, "files");
            if (!Directory.Exists(exportGGFileDir))
            {
                Directory.CreateDirectory(exportGGFileDir);
            }
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> listRow = listSheet.GetRow(i);

                bool giveUp = "Y".Equals(listRow[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string url = listRow[SysConfig.DetailPageUrlFieldName];
                    string pdfFilePath = this.RunPage.GetFilePath(url, pageSourceDir);
                    string fileName = Path.GetFileName(pdfFilePath);
                    //string pdfFilePath = this.RunPage.GetFilePath(url, pageSourceDir);
                    //string newFilePath = this.RunPage.GetFilePath(listRow["pinyin"] + "_" + listRow["code"] + "_" + listRow["zwjc"] + "_" + listRow["announcementTitle"] + "_" + url, exportGGFileDir);
                    //string newFileName = Path.GetFileName(newFilePath);

                    //if (!File.Exists(newFilePath))
                    //{
                        //File.Copy(pdfFilePath, newFilePath);
                    //}
                    Dictionary<string, object> resultRow = new Dictionary<string, object>();
                    resultRow.Add("url", url);
                    resultRow.Add("fileName", fileName);
                    resultRow.Add("简写", listRow["pinyin"]);
                    resultRow.Add("名称", listRow["zwjc"]);
                    resultRow.Add("编码", listRow["code"]);
                    resultRow.Add("标题", listRow["announcementTitle"]);
                    resultEW.AddRow(resultRow);
                }
            }
            resultEW.SaveToDisk();
        } 
    }
}