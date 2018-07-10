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

namespace NetDataAccess.Extended.GXP
{
    /// <summary>
    /// HenanQiPingDangan
    /// </summary>
    public class HenanQiPingDangan : ExternalRunWebPage
    {
        public override byte[] GetRequestData_BeforeSendRequest(string pageUrl, Dictionary<string, string> listRow, Encoding encoding)
        {
            return encoding.GetBytes(listRow["data"]);
        }

        public override void WebRequestHtml_BeforeSendRequest(string pageUrl, Dictionary<string, string> listRow, Base.Web.NDAWebClient client)
        {
            client.Headers.Add("content-type", "application/x-www-form-urlencoded");
        }

        public override bool AfterAllGrab(IListSheet listSheet)
        {
            return GenerateDAGLJL(listSheet) && GenerateCLXX(listSheet);
        }

        /// <summary>
        /// 生成档案管理记录
        /// </summary>
        /// <param name="listSheet"></param>
        /// <returns></returns>
        private bool GenerateDAGLJL(IListSheet listSheet)
        {
            bool succeed = true;
            string exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> dagljlColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                        "区划",
                        "录入单位",
                        "登记证编号", 
                        "车牌号码", 
                        "安装数量", 
                        "使用单位", 
                        "安装日期",
                        "登记日期",
                        "状态"});
            string dagljlPath = Path.Combine(exportDir, this.RunPage.Project.Name + "_档案管理记录.xlsx");
            Dictionary<string, string> columnFormats = new Dictionary<string, string>(); 
            ExcelWriter cityReportEW = new ExcelWriter(dagljlPath, "List", dagljlColumnDic, columnFormats);


            int detailUrlColumnIndex = this.RunPage.ColumnNameToIndex["detailPageUrl"];
            Dictionary<string, string> codeDateToNull = new Dictionary<string, string>(); 

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"]; 
                string localFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir);
                TextReader tr = null;

                try
                {
                    tr = new StreamReader(localFilePath, Encoding.GetEncoding(((Proj_Detail_SingleLine)this.RunPage.Project.DetailGrabInfoObject).Encoding));
                    string webPageHtml = tr.ReadToEnd();

                    HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                    htmlDoc.LoadHtml(webPageHtml);

                    HtmlNodeCollection listTrNodeList = htmlDoc.DocumentNode.SelectNodes("//table[@class=\"GridView\"]/tr");
                    if (listTrNodeList.Count > 1)
                    {
                        for (int j = 1; j < listTrNodeList.Count; j++)
                        {
                            HtmlNode listTrNode = listTrNodeList[j];

                            HtmlNodeCollection vNodeList = listTrNode.SelectNodes("./td");
                            Dictionary<string, object> reportInfo = new Dictionary<string, object>();

                            reportInfo.Add("区划", vNodeList[1].InnerText.Trim());
                            reportInfo.Add("录入单位", vNodeList[2].InnerText.Trim());
                            reportInfo.Add("登记证编号", vNodeList[3].InnerText.Trim());
                            reportInfo.Add("车牌号码", vNodeList[4].InnerText.Trim());
                            reportInfo.Add("安装数量", vNodeList[5].InnerText.Trim());
                            reportInfo.Add("使用单位", vNodeList[6].InnerText.Trim());

                            reportInfo.Add("安装日期", vNodeList[7].InnerText.Trim());
                            reportInfo.Add("登记日期", vNodeList[8].InnerText.Trim());
                            reportInfo.Add("状态", vNodeList[9].InnerText.Trim());
                            cityReportEW.AddRow(reportInfo);
                        }
                    }
                }
                catch (Exception ex)
                {
                    if (tr != null)
                    {
                        tr.Dispose();
                        tr = null;
                    }
                    this.RunPage.InvokeAppendLogText("读取出错. " + ex.Message + " LocalPath = " + localFilePath, LogLevelType.Error, true);
                }
            }
            cityReportEW.SaveToDisk();
            return succeed;
        }
        /// <summary>
        /// 生成车辆信息抓取URL列表
        /// </summary>
        /// <param name="listSheet"></param>
        /// <returns></returns>
        private bool GenerateCLXX(IListSheet listSheet)
        {
            bool succeed = true;
            string exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> clxxColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                    "detailPageUrl",
                    "detailPageName", 
                    "cookie",
                    "grabStatus", 
                    "giveUpGrab"});
            string clxxPath = Path.Combine(exportDir, this.RunPage.Project.Name + "_车辆信息.xlsx");
            ExcelWriter clxxEW = new ExcelWriter(clxxPath, "List", clxxColumnDic, null);


            int detailUrlColumnIndex = this.RunPage.ColumnNameToIndex["detailPageUrl"];
            Dictionary<string, string> rIdToNull = new Dictionary<string, string>(); 

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"];
                string cookie = row["cookie"]; 
                string localFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir);
                TextReader tr = null;

                try
                {
                    tr = new StreamReader(localFilePath, Encoding.GetEncoding(((Proj_Detail_SingleLine)this.RunPage.Project.DetailGrabInfoObject).Encoding));
                    string webPageHtml = tr.ReadToEnd();

                    HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                    htmlDoc.LoadHtml(webPageHtml);

                    HtmlNodeCollection listTrNodeList = htmlDoc.DocumentNode.SelectNodes("//table[@class=\"GridView\"]/tr");
                    if (listTrNodeList.Count > 1)
                    {
                        for (int j = 1; j < listTrNodeList.Count; j++)
                        {
                            HtmlNode listTrNode = listTrNodeList[j];

                            HtmlNodeCollection vNodeList = listTrNode.SelectNodes("./td");
                            Dictionary<string, object> reportInfo = new Dictionary<string, object>();
                            string clickUrl = vNodeList[3].SelectSingleNode("./span/a").GetAttributeValue("onclick", "");
                            string rId = clickUrl.Substring(clickUrl.IndexOf("=") + 1, clickUrl.LastIndexOf("'") - clickUrl.IndexOf("=") - 1);
                            if (!rIdToNull.ContainsKey(rId))
                            {
                                string pageUrl = "http://218.56.62.250/hnts/VehicleGas/VehicleGasView.aspx?RegId=" + rId;
                                rIdToNull.Add(rId, ""); 
                                reportInfo.Add("detailPageUrl", pageUrl);
                                reportInfo.Add("detailPageName", rId);
                                reportInfo.Add("cookie", cookie);
                                clxxEW.AddRow(reportInfo);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    if (tr != null)
                    {
                        tr.Dispose();
                        tr = null;
                    }
                    this.RunPage.InvokeAppendLogText("读取出错. " + ex.Message + " LocalPath = " + localFilePath, LogLevelType.Error, true);
                }
            }
            clxxEW.SaveToDisk();
            return succeed;
        }         
    }

}