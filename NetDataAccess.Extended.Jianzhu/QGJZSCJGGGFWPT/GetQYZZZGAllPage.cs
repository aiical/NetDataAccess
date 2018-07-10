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
    public class GetQYZZZGAllPage : ExternalRunWebPage
    {
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
            resultColumnDic.Add("资质类别", 1);
            resultColumnDic.Add("资质证书号", 2);
            resultColumnDic.Add("资质名称", 3);
            resultColumnDic.Add("发证日期", 4);
            resultColumnDic.Add("证件有效期", 5);
            resultColumnDic.Add("发证机关", 6);

            string resultFilePath = Path.Combine(exportDir, "企业数据_企业资质资格.xlsx");
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
                                HtmlNode indexNode = tdNodeList[0];
                                if (indexNode.GetAttributeValue("data-header", "") == "序号")
                                {
                                    try
                                    {
                                        Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                        f2vs.Add("CompanyId", companyId);
                                        f2vs.Add("资质类别", tdNodeList.Count < 2 ? "" : tdNodeList[1].InnerText.Trim());
                                        f2vs.Add("资质证书号", tdNodeList.Count < 3 ? "" : tdNodeList[2].InnerText.Trim());
                                        f2vs.Add("资质名称", tdNodeList.Count < 4 ? "" : tdNodeList[3].InnerText.Trim());
                                        f2vs.Add("发证日期", tdNodeList.Count < 5 ? "" : tdNodeList[4].InnerText.Trim());
                                        f2vs.Add("证件有效期", tdNodeList.Count < 6 ? "" : tdNodeList[5].InnerText.Trim());
                                        f2vs.Add("发证机关", tdNodeList.Count < 7 ? "" : tdNodeList[6].InnerText.Trim());
                                        cw.AddRow(f2vs);
                                    }
                                    catch (Exception ex)
                                    {
                                        throw ex;
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