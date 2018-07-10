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
using System.Web;
using NetDataAccess.Base.Reader;

namespace NetDataAccess.Extended.Xingzheng.QuHua
{
    public class GetQuHuaPage : ExternalRunWebPage
    { 
         
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            HtmlAgilityPack.HtmlDocument pageHtmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, 0);
            HtmlNodeCollection itemNodeList = pageHtmlDoc.DocumentNode.SelectNodes("//p[@class=\"MsoNormal\"]");

            ExcelWriter ew = this.GetExcelWriter();
            for (int i = 0; i < itemNodeList.Count; i++)
            { 
                HtmlNode itemNode = itemNodeList[i];
                try
                {
                    string[] strParts =itemNode.InnerText.Split(new string[] { }, StringSplitOptions.RemoveEmptyEntries);
                    /*
                    HtmlNodeCollection bSpanNodeList = itemNode.SelectNodes("./b/span");
                    if (bSpanNodeList != null && bSpanNodeList.Count > 0)
                    {
                        code = bSpanNodeList[0].InnerText.Trim();
                        name = bSpanNodeList[1].InnerText.Trim();
                    }
                    else
                    {
                        HtmlNodeCollection spanNodeList = itemNode.SelectNodes("./span");
                        code = spanNodeList[1].InnerText.Trim();
                        name = spanNodeList[2].InnerText.Trim();
                    }*/

                    Dictionary<string, string> f2vs = new Dictionary<string, string>();
                    f2vs.Add("code", CommonUtil.HtmlDecode(strParts[0]).Trim());
                    f2vs.Add("name", CommonUtil.HtmlDecode(strParts[1]).Trim());
                    ew.AddRow(f2vs);
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
            ew.SaveToDisk();

            return true;
        }
        private ExcelWriter GetExcelWriter()
        {
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("code", 0);
            resultColumnDic.Add("name", 1); 

            string resultFilePath = Path.Combine(exportDir, "行政区划.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic);
            return resultEW;
        }
    }
}