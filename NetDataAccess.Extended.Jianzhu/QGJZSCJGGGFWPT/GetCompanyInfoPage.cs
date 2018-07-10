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
using NetDataAccess.Base.DataTransform.Address;

namespace NetDataAccess.Extended.Jianzhu.QGJZSCJGGGFWPT
{
    public class GetCompanyInfoPage : ExternalRunWebPage
    { 

        public override bool AfterAllGrab(IListSheet listSheet)
        { 
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("CompanyId", 0);
            resultColumnDic.Add("企业名称", 1);
            resultColumnDic.Add("统一社会信用代码", 2);
            resultColumnDic.Add("企业法定代表人", 3);
            resultColumnDic.Add("企业登记注册类型", 4);
            resultColumnDic.Add("企业注册属地", 5);
            resultColumnDic.Add("企业经营地址", 6); 
            string resultFilePath = Path.Combine(exportDir, "企业数据_企业基本信息.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;
            Dictionary<string, string> companyDic = new Dictionary<string, string>();

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string url = row[detailPageUrlColumnName];
                    string certNo = row["certNo"];
                    string companyId = row["companyId"];

                    HtmlAgilityPack.HtmlDocument pageHtmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                    HtmlNode companyNameNode = pageHtmlDoc.DocumentNode.SelectSingleNode("//div[@class=\"user_info spmtop\"]");
                    string companyName = companyNameNode.InnerText.ToString();
                    string tyshxydm = "";
                    string qyfddbr = "";
                    string qydjzclx = "";
                    string qyzcsd = "";
                    string qyjydz = "";

                    HtmlNodeCollection companyFieldNodes = pageHtmlDoc.DocumentNode.SelectNodes("//table[@class=\"pro_table_box datas_table\"]/tbody/tr/td");
                    foreach (HtmlNode companyFieldNode in companyFieldNodes)
                    {
                        string dataHeader = companyFieldNode.GetAttributeValue("data-header", "");
                        switch (dataHeader)
                        {
                            case "统一社会信用代码":
                                tyshxydm = companyFieldNode.InnerText.Trim();
                                break;
                            case "企业法定代表人":
                                qyfddbr = companyFieldNode.InnerText.Trim();
                                break;
                            case "企业登记注册类型":
                                qydjzclx = companyFieldNode.InnerText.Trim();
                                break;
                            case "企业注册属地":
                                qyzcsd = companyFieldNode.InnerText.Trim();
                                break;
                            case "企业经营地址":
                                qyjydz = companyFieldNode.InnerText.Trim();
                                break;
                        }
                    }
                    Dictionary<string, string> f2vs = new Dictionary<string, string>();
                    f2vs.Add("CompanyId", companyId);
                    f2vs.Add("企业名称", companyName);
                    f2vs.Add("统一社会信用代码", tyshxydm);
                    f2vs.Add("企业法定代表人", qyfddbr);
                    f2vs.Add("企业登记注册类型", qydjzclx);
                    f2vs.Add("企业注册属地", qyzcsd);
                    f2vs.Add("企业经营地址", qyjydz); 

                    resultEW.AddRow(f2vs);
                }
            }

            resultEW.SaveToDisk();

            return true;
        }
    }
}