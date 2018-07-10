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

namespace NetDataAccess.Extended.Jianzhu.QGJZSCJGGGFWPT
{
    public class GetQualList : ExternalRunWebPage
    {
        public override byte[] GetRequestData_BeforeSendRequest(string pageUrl, Dictionary<string, string> listRow, Encoding encoding)
        {
            string pageIndex = listRow["pageIndex"];
            string qualCount = listRow["qualCount"];
            string data = "%24total=" + qualCount.ToString() + "&%24reload=0&%24pg=" + pageIndex.ToString() + "&%24pgsz=10";
            return encoding.GetBytes(data);
        }
        public override void WebRequestHtml_BeforeSendRequest(string pageUrl, Dictionary<string, string> listRow, Base.Web.NDAWebClient client)
        {
            client.Headers.Add("content-type", "application/x-www-form-urlencoded");
        }


        #region 获取资质
        public override bool AfterAllGrab(IListSheet listSheet)
        {  
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
             
            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{ 
                "aptCode",
                "aptScope"});
             
            string exportDir = this.RunPage.GetExportDir();

            string resultFilePath = Path.Combine(exportDir, "企业数据_资质.xlsx");
             
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic); 
             
            GetQuals(listSheet, pageSourceDir, resultEW);

            //保存到硬盘
            resultEW.SaveToDisk();

            return true;
        }
        #endregion

        #region GetQuals
        /// <summary>
        /// GetCities
        /// </summary>
        /// <param name="listSheet"></param>
        /// <param name="pageSourceDir"></param>
        /// <param name="resultEW"></param>
        private void GetQuals(IListSheet listSheet, string pageSourceDir, ExcelWriter resultEW)
        {
            Dictionary<string, string> codeDic = new Dictionary<string, string>();
            for (int i = 0; i < listSheet.RowCount; i++)
            { 
                string pageUrl = listSheet.PageUrlList[i]; 
                HtmlAgilityPack.HtmlDocument pageHtmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                HtmlNodeCollection allQualNodes = pageHtmlDoc.DocumentNode.SelectNodes("//input[@class=\"icheck\"]");
                if (allQualNodes != null)
                {
                    for (int j = 0; j < allQualNodes.Count; j++)
                    {
                        String jsonText = allQualNodes[j].GetAttributeValue("value", "");
                        JObject rootJo = JObject.Parse(jsonText);
                        string aptCode = (rootJo.SelectToken("apt_code") as JValue).ToString().Trim();
                        string aptScope = (rootJo.SelectToken("apt_scope") as JValue).ToString().Trim();
                        if (!codeDic.ContainsKey(aptCode))
                        {
                            codeDic.Add(aptCode, null);
                            Dictionary<string, string> f2vs = new Dictionary<string, string>();
                            f2vs.Add("aptCode", aptCode);
                            f2vs.Add("aptScope", aptScope);
                            resultEW.AddRow(f2vs);
                        }
                    }
                }
            }
        }
        #endregion
    }
}