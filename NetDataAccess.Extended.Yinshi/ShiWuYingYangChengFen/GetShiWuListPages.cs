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

namespace NetDataAccess.Extended.Yinshi.ShiWuYingYangChengFen
{
    public class GetShiWuListPages : ExternalRunWebPage
    { 
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            try
            {
                this.GetList(listSheet); 
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return true;
        }

        private void GetList(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("所属类别", 5);
            resultColumnDic.Add("名称", 6);
            string resultFilePath = Path.Combine(exportDir, "食物营养成分详情页.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null); 
            
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"];
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                { 
                    try
                    {
                        HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                        HtmlNodeCollection itemNodes = htmlDoc.DocumentNode.SelectNodes("//ul[@class=\"lst3\"]/li/a");

                        foreach (HtmlNode itemNode in itemNodes)
                        {
                            string href = itemNode.GetAttributeValue("href", "");
                            string name = itemNode.InnerText.Trim();

                            Dictionary<string, string> f2vs = new Dictionary<string, string>();
                            f2vs.Add("detailPageUrl", "http://yingyang.118cha.com/" + href);
                            f2vs.Add("detailPageName", href);
                            f2vs.Add("所属类别", row["分类名称"]);
                            f2vs.Add("名称", name);
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
    }
}