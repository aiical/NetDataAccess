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

namespace NetDataAccess.Extended.GuPiao
{
    public class GetGongSiRenYuanDetailPage : ExternalRunWebPage
    { 
        public override bool AfterAllGrab(IListSheet listSheet)
        { 
            this.GetRenYuanDetailInfos(listSheet);
            return true;
        }

        private void GetRenYuanDetailInfos(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("公司代码", 0);
            resultColumnDic.Add("姓名", 1);
            resultColumnDic.Add("性别", 2);
            resultColumnDic.Add("出生日期", 3);
            resultColumnDic.Add("学历", 4);
            resultColumnDic.Add("国籍", 5);
            resultColumnDic.Add("简介", 6);
            string resultFilePath = Path.Combine(exportDir, "上市公司高管人员信息.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i); 
                string detailUrl = row["detailPageUrl"];
                string gsdm = row["公司代码"];
                string xm = row["姓名"];

                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    HtmlAgilityPack.HtmlDocument pageHtmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i, Encoding.GetEncoding("gb2312"));

                    try
                    {
                        HtmlNodeCollection tdNodes = pageHtmlDoc.DocumentNode.SelectNodes("//table[@id=\"Table1\"]/tbody/tr/td");

                        Dictionary<string, string> f2vs = new Dictionary<string, string>();
                        f2vs.Add("公司代码", gsdm);
                        f2vs.Add("姓名", xm);
                        f2vs.Add("性别", CommonUtil.HtmlDecode(tdNodes[1].InnerText).Trim());
                        f2vs.Add("出生日期", CommonUtil.HtmlDecode(tdNodes[2].InnerText).Trim());
                        f2vs.Add("学历", CommonUtil.HtmlDecode(tdNodes[3].InnerText).Trim());
                        f2vs.Add("国籍", CommonUtil.HtmlDecode(tdNodes[4].InnerText).Trim());
                        f2vs.Add("简介", CommonUtil.HtmlDecode(tdNodes[6].InnerText).Trim());
                        resultEW.AddRow(f2vs);
                    }
                    catch (Exception ex)
                    {
                        this.RunPage.InvokeAppendLogText(ex.Message + ", detailUrl = " + detailUrl, LogLevelType.Error, true);
                        throw ex;
                    }
                }
            }
            resultEW.SaveToDisk();
        }          
    }
}