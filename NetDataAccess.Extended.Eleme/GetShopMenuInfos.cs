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
using NetDataAccess.Base.Reader;

namespace NetDataAccess.Extended.Eleme
{
    public class GetShopMenuInfos : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            return true;
        }

        public override void WebRequestHtml_BeforeSendRequest(string pageUrl, Dictionary<string, string> listRow, Base.Web.NDAWebClient client)
        {
            base.WebRequestHtml_BeforeSendRequest(pageUrl, listRow, client);
            string xShard = listRow["xShard"];

            client.Headers.Add("x-shard", xShard);
            client.Headers.Add("content-type", "application/json");
        }
        public override byte[] GetRequestData_BeforeSendRequest(string pageUrl, Dictionary<string, string> listRow, Encoding encoding)
        {
            string dataText = listRow["dataText"];
            byte[] dataArray = Encoding.UTF8.GetBytes(dataText);
            return dataArray;
        } 

        private ExcelWriter CreateDetailFileWriter(int fileIndex)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            string resultFilePath = Path.Combine(exportDir, "饿了么_店铺详情页_" + fileIndex.ToString() + ".xlsx");
            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("address", 5);
            resultColumnDic.Add("description", 6);
            resultColumnDic.Add("id", 7);
            resultColumnDic.Add("latitude", 8);
            resultColumnDic.Add("longitude", 9);
            resultColumnDic.Add("name", 10);
            resultColumnDic.Add("phone", 11);
            resultColumnDic.Add("promotion_info", 12);
            resultColumnDic.Add("searchLat", 13);
            resultColumnDic.Add("searchLng", 14);
            resultColumnDic.Add("elemeCity", 15);
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }
    }
}