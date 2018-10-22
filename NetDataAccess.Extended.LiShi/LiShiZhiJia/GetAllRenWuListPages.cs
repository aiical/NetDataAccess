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

namespace NetDataAccess.Extended.LiShi.LiShiZhiJia
{
    public class GetAllRenWuListPages : ExternalRunWebPage
    { 
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            this.GetListPageUrls(listSheet); 
            return true;
        }

        private void GetListPageUrls(IListSheet listSheet)
        {
            ExcelWriter ew = this.CreateWriter();

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> listRow = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(listRow[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    try
                    {
                        HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);

                        HtmlNodeCollection linkNodes = htmlDoc.DocumentNode.SelectNodes("//div[@class=\"cont\"]/a");
                        for (int j = 0; j < linkNodes.Count; j++)
                        {
                            HtmlNode linkNode = linkNodes[j];
                            string url = "http://www.lszj.com" + linkNode.GetAttributeValue("href", "");
                            string name = CommonUtil.HtmlDecode(linkNode.InnerText).Trim();
                            Dictionary<string, string> row = new Dictionary<string, string>();
                            row.Add("detailPageUrl", url);
                            row.Add("detailPageName", url);
                            row.Add("name", name);
                            ew.AddRow(row);
                        }
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
            }
            ew.SaveToDisk();
        }

        private ExcelWriter CreateWriter()
        {
            String exportDir = this.RunPage.GetExportDir(); 
            string resultFilePath = Path.Combine(exportDir, "历史_历史之家_人物详情页.xlsx");

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("name", 5);
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }
         
    }
}