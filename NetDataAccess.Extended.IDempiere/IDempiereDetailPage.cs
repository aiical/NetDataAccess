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
using System.Web; 

namespace NetDataAccess.Extended.IDempiere
{
    public class IDempiereDetailPage : CustomProgramBase
    { 
        public bool Run(string parameters, IListSheet listSheet )
        {
            try
            {
                return this.GenerateNewPage(listSheet); 
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        private bool GenerateNewPage(IListSheet listSheet)
        {
            bool succeed = true;
            string exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            string localHtmlDir = Path.Combine(exportDir, "LocalHtml");
            if (!Directory.Exists(localHtmlDir))
            {
                Directory.CreateDirectory(localHtmlDir);
            }
            for (int i = 0; i < listSheet.RowCount ; i++)
            {
                Dictionary<string,string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string url = row[SysConfig.DetailPageUrlFieldName];
                    string name = row[SysConfig.DetailPageNameFieldName];
                    string localFilePath = this.RunPage.GetFilePath(url, pageSourceDir);
                    HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                    HtmlNode tableNode = htmlDoc.DocumentNode.SelectSingleNode("//body/table[1]");
                    tableNode.Attributes["border"].Value = "1";
                    string destFilePath = Path.Combine(localHtmlDir, name + ".html");
                    htmlDoc.Save(destFilePath); 
                }
            }  
            return succeed;
        } 
    }
}