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

namespace NetDataAccess.Extended.Jianzhu.QGJZSCJGGGFWPT
{
    public class GetQYZZZGFirstPage : ExternalRunWebPage
    { 
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            try
            {
                return this.GetAllQYZZZGPageUrls(listSheet); 
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private ExcelWriter GetQYZZZGExcelWriter()
        { 
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("companyId", 5);
            resultColumnDic.Add("pageIndex", 6);

            string resultFilePath = Path.Combine(exportDir, "企业数据_企业资质资格全部页面.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }

        private bool GetAllQYZZZGPageUrls(IListSheet listSheet)
        { 
            ExcelWriter qyzzzgEW = this.GetQYZZZGExcelWriter(); 

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;
            Dictionary<string, string> companyDic = new Dictionary<string, string>();

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailPageUrl = row[SysConfig.DetailPageUrlFieldName];
                string detailPageName = row[SysConfig.DetailPageNameFieldName];
                string companyId = row["companyId"];

                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    HtmlAgilityPack.HtmlDocument pageHtmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);

                    //都包含第一页
                    Dictionary<string, string> f2vs = new Dictionary<string, string>();
                    f2vs.Add("detailPageUrl", detailPageUrl);
                    f2vs.Add("detailPageName", detailPageName);
                    f2vs.Add("companyId", companyId);
                    f2vs.Add("pageIndex", "1");
                    qyzzzgEW.AddRow(f2vs);

                    HtmlNode pageNode = pageHtmlDoc.DocumentNode.SelectSingleNode("//a[@sf=\"pagebar\"]");
                    if (pageNode != null)
                    {
                        string pageData = pageNode.GetAttributeValue("sf:data", "");
                        if (pageData.Length == 0)
                        {
                            throw new Exception("获取分页信息错误. companyId = " + companyId);
                        }
                        else
                        {
                            int pcBeginIndex = pageData.IndexOf(",pc:");
                            if (pcBeginIndex < 0)
                            {
                                throw new Exception("获取分页信息错误. companyId = " + companyId);
                            }
                            else
                            {
                                int pcEndIndex = pageData.IndexOf(",", pcBeginIndex + 4);
                                int pageCount = int.Parse(pageData.Substring(pcBeginIndex + 4, pcEndIndex - pcBeginIndex - 4));
                                for (int pIndex = 2; pIndex <= pageCount; pIndex++)
                                {
                                    Dictionary<string, string> otherPageF2vs = new Dictionary<string, string>();
                                    otherPageF2vs.Add("detailPageUrl", detailPageUrl + "&$pg=" + pIndex.ToString());
                                    otherPageF2vs.Add("detailPageName", detailPageName + "_" + pIndex.ToString());
                                    otherPageF2vs.Add("companyId", companyId);
                                    otherPageF2vs.Add("pageIndex", pIndex.ToString());
                                    qyzzzgEW.AddRow(otherPageF2vs);
                                }
                            }
                        } 
                    } 
                }
            }
             
            qyzzzgEW.SaveToDisk(); 

            return true;
        }
         
    }
}