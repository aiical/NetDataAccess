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
    public class GetQYGCXMAllPage : ExternalRunWebPage
    {
        public override byte[] GetRequestData_BeforeSendRequest(string pageUrl, Dictionary<string, string> listRow, Encoding encoding)
        {
            string formData = listRow["formData"];
            if (formData != null && formData.Length > 0)
            {
                return encoding.GetBytes(formData);
            }
            else
            {
                return null;
            }
        } 
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            try
            {
                this.GetAllInfos(listSheet);
                return true; 
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private CsvWriter GetCsvExcelWriter()
        {
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("CompanyId", 0);
            resultColumnDic.Add("项目编码", 1);
            resultColumnDic.Add("项目名称", 2);
            resultColumnDic.Add("项目属地", 3);
            resultColumnDic.Add("项目类别", 4);
            resultColumnDic.Add("建设单位", 5); 

            string resultFilePath = Path.Combine(exportDir, "企业数据_企业工程项目.xlsx");
            CsvWriter resultEW = new CsvWriter(resultFilePath, resultColumnDic);
            return resultEW;
        }
        private void GetAllInfos(IListSheet listSheet)
        {
            CsvWriter cw = this.GetCsvExcelWriter(); 

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
                      
                    HtmlNodeCollection trNodeList = pageHtmlDoc.DocumentNode.SelectNodes("//table/tbody/tr");
                    if (trNodeList != null)
                    {
                        for (int j = 0; j < trNodeList.Count; j++)
                        {
                            try
                            {
                                HtmlNode trNode = trNodeList[j];
                                HtmlNodeCollection tdNodeList = trNode.SelectNodes("./td");
                                if (tdNodeList != null && tdNodeList.Count > 0)
                                {
                                    HtmlNode indexNode = tdNodeList[0];
                                    if (indexNode.GetAttributeValue("data-header", "") == "序号")
                                    {
                                        try
                                        {
                                            Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                            f2vs.Add("CompanyId", companyId);
                                            f2vs.Add("项目编码", tdNodeList.Count < 2 ? "" : tdNodeList[1].InnerText.Trim());
                                            f2vs.Add("项目名称", tdNodeList.Count < 3 ? "" : tdNodeList[2].InnerText.Trim());
                                            f2vs.Add("项目属地", tdNodeList.Count < 4 ? "" : tdNodeList[3].InnerText.Trim());
                                            f2vs.Add("项目类别", tdNodeList.Count < 5 ? "" : tdNodeList[4].InnerText.Trim());
                                            f2vs.Add("建设单位", tdNodeList.Count < 6 ? "" : tdNodeList[5].InnerText.Trim());
                                            cw.AddRow(f2vs);
                                        }
                                        catch (Exception ex)
                                        {
                                            throw ex;
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
                }
            }
             
            cw.SaveToDisk();  
        }         
    }
}