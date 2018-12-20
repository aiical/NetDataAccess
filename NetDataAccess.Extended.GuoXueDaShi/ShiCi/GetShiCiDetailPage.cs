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

namespace NetDataAccess.Extended.GuoXueDaShi.ShiCi
{
    public class GetShiCiDetailPage : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            this.GetShiCiInfos(listSheet);
            return true;
        }

        private void GetShiCiInfos(IListSheet listSheet)
        {
            string sourceDir = this.RunPage.GetDetailSourceFileDir();
            ExcelWriter resultEW = this.CreateRenWuResultWriter();
            Dictionary<string, bool> pageUrlDic = new Dictionary<string, bool>();
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> listRow = listSheet.GetRow(i);
                HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);

                HtmlNode mainInfoNode = htmlDoc.DocumentNode.SelectSingleNode("//div[@class=\"info_txt2 clearfix\"]"); 
                HtmlNode infoNode = mainInfoNode.SelectSingleNode("./p");
                StringBuilder infoStrBuilder = new StringBuilder();
                foreach (HtmlNode childNode in infoNode.ChildNodes)
                {
                    if (childNode.Name.ToLower() == "br")
                    {
                        infoStrBuilder.AppendLine("\r\n");
                    }
                    else
                    {
                        string partStr = CommonUtil.HtmlDecode(childNode.InnerText).Trim();
                        if (partStr != null && partStr.Length > 0)
                        {
                            infoStrBuilder.AppendLine(partStr);
                        }
                    }

                    Dictionary<string, string> resultRow = new Dictionary<string, string>();
                    resultRow.Add("人物", listRow["renWu"]);
                    resultRow.Add("时代", listRow["shiDai"]);
                    resultRow.Add("诗词名称", listRow["shiCi"]);
                    resultRow.Add("内容", infoStrBuilder.ToString().Trim());
                    resultRow.Add("url", listRow[SysConfig.DetailPageUrlFieldName]);
                    resultEW.AddRow(resultRow);
                }
            }

            resultEW.SaveToDisk();
        }

        private ExcelWriter CreateRenWuResultWriter()
        {
            String exportDir = this.RunPage.GetExportDir();
            string resultFilePath = Path.Combine(exportDir, "国学大师_诗词_诗词信息.xlsx");

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("人物", 0);
            resultColumnDic.Add("时代", 1);
            resultColumnDic.Add("诗词名称", 2);
            resultColumnDic.Add("内容", 3);
            resultColumnDic.Add("url", 4); 
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }         
    }
}