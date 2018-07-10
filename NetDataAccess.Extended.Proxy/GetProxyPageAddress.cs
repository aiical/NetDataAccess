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
using NetDataAccess.Base.Web;
using System.Net;

namespace NetDataAccess.Extended.Proxy
{
    public class GetProxyPageAddress : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        { 
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("ip", 0);
            resultColumnDic.Add("port", 1);
            resultColumnDic.Add("user", 2);
            resultColumnDic.Add("pwd", 3);
            resultColumnDic.Add("usable", 4);
            resultColumnDic.Add("fromSiteName", 5);
            resultColumnDic.Add("timespan", 6);
            resultColumnDic.Add("address", 6);
            string resultFilePath = Path.Combine(exportDir, "Proxy.xlsx");

            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "Proxy", resultColumnDic, null);

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string detailPageUrl = row["detailPageUrl"];
                    string detailPageName = row["detailPageName"];
                    HtmlAgilityPack.HtmlDocument pageHtmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i, Encoding.GetEncoding("gb2312"));
                    HtmlNode addressNode = pageHtmlDoc.DocumentNode.SelectSingleNode("//ul[@class=\"ul1\"]/li[1]");
                    string info = addressNode.InnerText;
                    int fromIndex = info.IndexOf("：")+ 1;
                    int endIndex = info.LastIndexOf("  ");
                    string address = fromIndex <= 0 || endIndex < 0 ? "" : info.Substring(fromIndex, endIndex - fromIndex);
                    Dictionary<string, string> r = new Dictionary<string, string>();
                    r.Add("ip", row["ip"]);
                    r.Add("port", row["port"]);
                    r.Add("user", row["user"]);
                    r.Add("pwd", row["pwd"]);
                    r.Add("usable", row["usable"]);
                    r.Add("fromSiteName", row["fromSiteName"]);
                    r.Add("timespan", row["timespan"]);
                    r.Add("address", address);
                    resultEW.AddRow(r);
                }
            }

            resultEW.SaveToDisk();

            return true;
        } 
    }
}