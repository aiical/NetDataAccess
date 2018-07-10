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
using NPOI.XSSF.UserModel;
using NetDataAccess.Base.DB; 

namespace NetDataAccess.Extended.IDempiere
{
    /// <summary>
    /// 处理页面
    /// </summary>
    public class ProcessPageUrl : CustomProgramBase
    { 
        public bool Run(string parameters, IListSheet listSheet )
        {
            try
            {
                return this.GeneratePageUrlList(listSheet); 
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        private bool GeneratePageUrlList(IListSheet listSheet)
        {
            bool succeed = true;
            string exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> pageListColumnDic = new Dictionary<string, int>(); 
            pageListColumnDic.Add("detailPageUrl", 0);
            pageListColumnDic.Add("detailPageName", 1);
            pageListColumnDic.Add("grabStatus", 2);
            pageListColumnDic.Add("giveUpGrab", 3);
            pageListColumnDic.Add("title", 4);
            pageListColumnDic.Add("creator", 5);
            pageListColumnDic.Add("createDate", 6);
            string pageListFilePath = Path.Combine(exportDir, this.RunPage.Project.Name + "_All.xlsx");
            ExcelWriter pageListEW = new ExcelWriter(pageListFilePath, "List", pageListColumnDic); 
            for (int i = 0; i < listSheet.RowCount ; i++)
            {
                Dictionary<string,string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                { 
                    string url = row[SysConfig.DetailPageUrlFieldName]; 
                    string localFilePath = this.RunPage.GetFilePath(url, pageSourceDir);
                    TextReader tr = null;

                    try
                    {
                        tr = new StreamReader(localFilePath, Encoding.Default, true);
                        string fileTxt = tr.ReadToEnd();

                        string[] pageLines = fileTxt.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
                        foreach (string pageLine in pageLines)
                        {
                            if (pageLine.Trim().Length != 0)
                            {
                                try
                                {
                                    string[] values = pageLine.Split(new string[] { "@@" }, StringSplitOptions.None);
                                    string title = values[1];
                                    string detailPageUrl = values[2];
                                    string detailPageName = detailPageUrl;
                                    string creator = values[3];
                                    string createDate = values[4];

                                    Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                    f2vs.Add("title", title);
                                    f2vs.Add("detailPageUrl", detailPageUrl);
                                    f2vs.Add("detailPageName", detailPageName);
                                    f2vs.Add("creator", creator);
                                    f2vs.Add("createDate", createDate);
                                    pageListEW.AddRow(f2vs);
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
                        if (tr != null)
                        {
                            tr.Dispose();
                            tr = null;
                        }
                        this.RunPage.InvokeAppendLogText("读取出错. " + ex.Message + " LocalPath = " + localFilePath, LogLevelType.Error, true);
                    }
                }
            }
            pageListEW.SaveToDisk();
            return succeed;
        } 
    }
}