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
    public class ProcessAtlassianPageUrl : CustomProgramBase
    { 
        public bool Run(string parameters, IListSheet listSheet )
        {
            try
            {
                return this.GenerateAtlassianPageUrlList(listSheet); 
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        private bool GenerateAtlassianPageUrlList(IListSheet listSheet)
        {
            bool succeed = true;
            string exportDir = this.RunPage.GetExportDir(); 

            Dictionary<string, int> pageListColumnDic = new Dictionary<string, int>(); 
            pageListColumnDic.Add("detailPageUrl", 0);
            pageListColumnDic.Add("detailPageName", 1);
            pageListColumnDic.Add("grabStatus", 2);
            pageListColumnDic.Add("giveUpGrab", 3);
            pageListColumnDic.Add("title", 4);
            string readDetailDir = this.RunPage.GetReadFileDir();
            string pageListFilePath = Path.Combine(exportDir, this.RunPage.Project.Name + "_All.xlsx");
            ExcelWriter pageListEW = new ExcelWriter(pageListFilePath, "List", pageListColumnDic); 
            for (int i = 0; i < listSheet.RowCount ; i++)
            {

                string pageUrl = listSheet.PageUrlList[i];
                string localReadFilePath = this.RunPage.GetReadFilePath(pageUrl, readDetailDir);
                List<Dictionary<string, string>> f2vsList = this.RunPage.ReadDetailFieldValueListFromFile(localReadFilePath);
                for (int j = 0; j < f2vsList.Count; j++)
                {
                    Dictionary<string, string> detailF2vs = f2vsList[j];
                    string title = detailF2vs["url"].Replace("/browse/","").Trim();
                    string detailPageUrl = "https://idempiere.atlassian.net/browse/" + title;

                    Dictionary<string, string> f2vs = new Dictionary<string, string>();
                    f2vs.Add("detailPageUrl", detailPageUrl);
                    f2vs.Add("detailPageName", detailPageUrl);
                    f2vs.Add("title", title);
                    pageListEW.AddRow(f2vs); 
                }
            }
            pageListEW.SaveToDisk();
            return succeed;
        } 
    }
}