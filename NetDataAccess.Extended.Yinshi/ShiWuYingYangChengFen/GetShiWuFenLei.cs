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
    public class GetShiWuFenLei : ExternalRunWebPage
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
            resultColumnDic.Add("分类名称", 5);
            string resultFilePath = Path.Combine(exportDir, "食物列表页.xlsx");
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
                        HtmlNodeCollection sectionNodes = htmlDoc.DocumentNode.SelectNodes("//p[@class=\"f14 l150\"]");

                        foreach (HtmlNode sectionNode in sectionNodes)
                        {
                            HtmlNode titleNode = sectionNode.SelectSingleNode("./strong");
                            if (titleNode != null && titleNode.InnerText.Trim() == "食物分类")
                            {
                                HtmlNodeCollection itemNodes = sectionNode.SelectNodes("./a");
                                foreach (HtmlNode itemNode in itemNodes)
                                {
                                    string href = itemNode.GetAttributeValue("href", "");
                                    string name = itemNode.InnerText.Trim(); 

                                    Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                    f2vs.Add("detailPageUrl", "http://yingyang.118cha.com/" + href);
                                    f2vs.Add("detailPageName", href);
                                    f2vs.Add("分类名称", name);
                                    resultEW.AddRow(f2vs);
                                }
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