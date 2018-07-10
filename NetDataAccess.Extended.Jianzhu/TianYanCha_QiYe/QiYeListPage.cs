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
using NetDataAccess.Base.UserAgent;

namespace NetDataAccess.Extended.Jianzhu.TianYanCha_QiYe
{
    public class QiYeListPage : ExternalRunWebPage
    {
        private UserAgents _RequestUserAgents = null;
        public override bool BeforeAllGrab()
        {
            this._RequestUserAgents = new UserAgents();
            this._RequestUserAgents.Load();
            return base.BeforeAllGrab();
        }

        public override void WebRequestHtml_BeforeSendRequest(string pageUrl, Dictionary<string, string> listRow, Base.Web.NDAWebClient client)
        {
            //string userAgent = this._RequestUserAgents.GetOnePcUserAgent();
            string userAgent = "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.113 Safari/537.36";
            client.Headers.Add("user-agent", userAgent);
            client.Headers.Add("cookie", "aliyungf_tc=AQAAAFK4lnkEvwcA2tbQPOb7I4znw5Ct; csrfToken=hv3pUW2vrQ14QiCYxgEXX0yw; TYCID=c2a7c83073e911e89329f931a4758275; undefined=c2a7c83073e911e89329f931a4758275; ssuid=5395707924; Hm_lvt_e92c8d65d92d534b0fc290df538b4758=1529430829; Hm_lpvt_e92c8d65d92d534b0fc290df538b4758=1529431243");
        }

        public override void CheckRequestCompleteFile(string webPageText, Dictionary<string, string> listRow)
        {
            HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
            htmlDoc.LoadHtml(webPageText);
            HtmlNodeCollection itemNodes = htmlDoc.DocumentNode.SelectNodes("//text[@class=\"tyc-num lh24\"]");
            if (itemNodes == null || itemNodes.Count == 0)
            {
                throw new Exception("没有找到对应的公司. companyName = " + listRow[SysConfig.DetailPageNameFieldName]);
            }
        }

        public override bool AfterAllGrab(IListSheet listSheet)
        {
            try
            {
                return this.GetAllListPageUrls(listSheet);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private ExcelWriter GetExcelWriter()
        {
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("detailPageUrl", 5);
            resultColumnDic.Add("detailPageName", 6);
            resultColumnDic.Add("cookie", 7);
            resultColumnDic.Add("grabStatus", 8);
            resultColumnDic.Add("giveUpGrab", 9);
            resultColumnDic.Add("formData", 10);

            string resultFilePath = Path.Combine(exportDir, "企业数据_企业工商信息.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        } 


        private bool GetAllListPageUrls(IListSheet listSheet)
        {
            int pageIndex = 1;
            bool needMoreFirstPage = false;
            {
                ExcelWriter ew = this.GetExcelWriter();

                string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;
                Dictionary<string, string> companyDic = new Dictionary<string, string>();
                for (int i = 0; i < listSheet.RowCount; i++)
                { 

                    Dictionary<string, string> row = listSheet.GetRow(i);
                    string detailPageUrl = row[SysConfig.DetailPageUrlFieldName];
                    string detailPageName = row[SysConfig.DetailPageNameFieldName];
                    string cookie = row[SysConfig.DetailPageCookieFieldName];

                    bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                    if (!giveUp)
                    {
                        HtmlAgilityPack.HtmlDocument pageHtmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                         
                    }
                }

                ew.SaveToDisk();
            }  

            return true;
        }
    }
}