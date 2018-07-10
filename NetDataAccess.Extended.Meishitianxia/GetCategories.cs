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

namespace NetDataAccess.Extended.Meishitianxia
{
    public class GetCategories : ExternalRunWebPage
    { 
        public override bool AfterAllGrab(IListSheet listSheet)
        { 
            this.GetListPageUrls(listSheet);
            return true;
        }

        private void GetListPageUrls(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("category", 5);
            resultColumnDic.Add("subCategory", 6);
            string resultFilePath = Path.Combine(exportDir, "美食天下_获取各小类菜谱列表页.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
             
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"]; 

                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    HtmlAgilityPack.HtmlDocument pageHtmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);

                    try
                    {
                        HtmlNodeCollection categoryDivList = pageHtmlDoc.DocumentNode.SelectNodes("//div[@class=\"category_sub clear\"]");

                        foreach (HtmlNode categoryDiv in categoryDivList)
                        {
                            HtmlNode categoryNameNode = categoryDiv.SelectSingleNode("./h3");
                            string categoryName = CommonUtil.HtmlDecode(categoryNameNode.InnerText).Trim();
                            HtmlNodeCollection subCategoryNodeList = categoryDiv.SelectNodes("./ul/li/a");
                            for (int j = 0; j < subCategoryNodeList.Count; j++)
                            {
                                HtmlNode subCategoryNode = subCategoryNodeList[j];
                                string subCategoryName = subCategoryNode.GetAttributeValue("title", "");
                                string subCategoryPageUrl = subCategoryNode.GetAttributeValue("href", "");
                                 
                                Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                f2vs.Add("detailPageUrl", subCategoryPageUrl);
                                f2vs.Add("detailPageName", subCategoryPageUrl);
                                f2vs.Add("category", categoryName);
                                f2vs.Add("subCategory", subCategoryName);

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