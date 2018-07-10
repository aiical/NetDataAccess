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
    public class GetQYGCXMFirstPage : ExternalRunWebPage
    { 
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            try
            {
                return this.GetAllQYGCXMPageUrls(listSheet); 
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private ExcelWriter GetQYGCXMExcelWriter()
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
            resultColumnDic.Add("formData", 7);

            string resultFilePath = Path.Combine(exportDir, "企业数据_企业工程项目全部页面.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }



        private bool GetAllQYGCXMPageUrls(IListSheet listSheet)
        {
            ExcelWriter qyzzzgEW = this.GetQYGCXMExcelWriter(); 

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
                    f2vs.Add("formData", "");
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
                            JObject rootJo = JObject.Parse(pageData.Substring(1, pageData.Length - 2));
                            string ps = rootJo.GetValue("ps").ToString();
                            string tt = rootJo.GetValue("tt").ToString();
                            string pc = rootJo.GetValue("pc").ToString();
                            int pageCount = int.Parse(pc);
                            for (int pIndex = 2; pIndex <= pageCount; pIndex++)
                            {
                                Dictionary<string, string> otherPageF2vs = new Dictionary<string, string>();
                                otherPageF2vs.Add("detailPageUrl", detailPageUrl + "?_=" + pIndex.ToString());
                                otherPageF2vs.Add("detailPageName", detailPageName + "_" + pIndex.ToString());
                                otherPageF2vs.Add("companyId", companyId);
                                otherPageF2vs.Add("pageIndex", pIndex.ToString());
                                otherPageF2vs.Add("formData", "%24total=" + tt + "&%24reload=0&%24pg=" + pIndex.ToString() + "&%24pgsz=" + ps);
                                qyzzzgEW.AddRow(otherPageF2vs);
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