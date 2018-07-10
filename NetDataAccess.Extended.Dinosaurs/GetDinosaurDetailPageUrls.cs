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

namespace NetDataAccess.Extended.Dinosaurs
{
    public class GetDinosaurDetailPageUrls : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            this.GetDetailPageUrls(listSheet);
            return true;
        }

        private void GetDetailPageUrls(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("name", 5);
            string resultFilePath = Path.Combine(exportDir, "www.nhm.ac.uk恐龙详情页.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            Dictionary<string, string> urlDic = new Dictionary<string, string>();
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"];
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string localFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir); 

                    try
                    {
                        HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                        htmlDoc.Load(localFilePath, Encoding.GetEncoding("utf-8"));

                        HtmlNodeCollection dinosaurNodes = htmlDoc.DocumentNode.SelectNodes("//ul[@class='dino-list dino-wrap']/li/a");

                        for (int j = 0; j < dinosaurNodes.Count; j++)
                        {
                            HtmlNode dinosaurNode = dinosaurNodes[j];
                            string name = dinosaurNode.InnerText.Trim();
                            string url = "http://www.nhm.ac.uk" + dinosaurNode.GetAttributeValue("href", "");
                            if (!urlDic.ContainsKey(url))
                            {
                                urlDic.Add(url, null);

                                Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                f2vs.Add("detailPageUrl", url);
                                f2vs.Add("detailPageName", url);
                                f2vs.Add("name", name); 
                                resultEW.AddRow(f2vs);
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