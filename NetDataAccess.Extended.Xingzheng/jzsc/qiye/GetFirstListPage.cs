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

namespace NetDataAccess.Extended.Jianzhu.jzsc.qiye
{
    public class GetFirstListPage : CustomProgramBase
    {
        public bool Run(string parameters, IListSheet listSheet)
        {
            return GetFirstListPageFromPage(parameters, listSheet);
        }
        private bool GetFirstListPageFromPage(string parameters, IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("pageNum", 5);
            resultColumnDic.Add("provinceName", 6);
            resultColumnDic.Add("provinceFullName", 7);
            resultColumnDic.Add("totalCount", 8);
            string resultFilePath = Path.Combine(exportDir, "企业数据_各省份全部.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string url = row[detailPageUrlColumnName];
                    string provinceId = row["regionId"];
                    string provinceName = row["regionName"];
                    string provinceFullName = row["regionFullName"];

                    HtmlAgilityPack.HtmlDocument pageHtmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                    string pageText = pageHtmlDoc.DocumentNode.SelectSingleNode("//form[@class=\"pagingform\"]").NextSibling.NextSibling.InnerText;
                    int totalStartIndex = pageText.IndexOf("\"$total\":") + 9;
                    int totalEndIndex = pageText.IndexOf(",", totalStartIndex);
                    string totalCount = pageText.Substring(totalStartIndex, totalEndIndex - totalStartIndex);
                    string detailPageUrl = "http://jzsc.mohurd.gov.cn/dataservice/query/comp/list?apt_code=&qy_fr_name=&%24total=" + totalCount + "&qy_reg_addr=" + provinceName + "&qy_reg_addr=" + provinceId + "&%24reload=0&qy_type=&qy_name=&%24pg=1&%24pgsz=" + totalCount + "&apt_scope=&apt_certno=";
                    Dictionary<string, string> f2vs = new Dictionary<string, string>();
                    f2vs.Add("detailPageUrl", detailPageUrl);
                    f2vs.Add("detailPageName", provinceId);
                    f2vs.Add("provinceId", provinceId);
                    f2vs.Add("provinceName", provinceName);
                    f2vs.Add("provinceFullName", provinceFullName);
                    f2vs.Add("totalCount", totalCount);
                    resultEW.AddRow(f2vs);
                }
            }

            resultEW.SaveToDisk();

            return true;
        }
        private void SaveRow(string infoName, string nodeHref, Dictionary<string, string> code2Names)
        {
            string[] infoPieces = nodeHref.Split(new string[] { "/" }, StringSplitOptions.RemoveEmptyEntries);
            string infoValue = infoPieces[infoPieces.Length - 1];
            if (!code2Names.ContainsKey(infoValue))
            {
                code2Names.Add(infoValue, infoName);
            }
        }
    }
}