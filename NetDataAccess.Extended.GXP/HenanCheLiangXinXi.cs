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
    /// HenanCheLiangXinXi
    /// </summary>
    public class HenanCheLiangXinXi : ExternalRunWebPage
    { 
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            return GenerateCLXXXX(listSheet) && GenerateCLQPLB(listSheet);
        }

        /// <summary>
        /// 生成车辆详细信息
        /// </summary>
        /// <param name="listSheet"></param>
        /// <returns></returns>
        private bool GenerateCLXXXX(IListSheet listSheet)
        {
            bool succeed = true;
            string exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> clxxColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                        "行政区划", 
                        "登记证编号",
                        "登记表编号",
                        "档案盒编号",
                        "车牌号码",
                        "发动机号", 
                        "车辆识别代码",  
                        "气瓶种类", 
                        "使用性质", 
                        "充装介质",
                        "使用单位",
                        "联系人",
                        "联系电话",
                        "起用日期",
                        "安装单位",
                        "安装日期",
                        "发证日期",
                        "发证机关"});
            string dagljlPath = Path.Combine(exportDir, this.RunPage.Project.Name + "_车辆详细信息.xlsx");
            Dictionary<string, string> columnFormats = new Dictionary<string, string>();
            ExcelWriter clxxEW = new ExcelWriter(dagljlPath, "List", clxxColumnDic, columnFormats);


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

                    HtmlNodeCollection listTrNodeList = htmlDoc.DocumentNode.SelectNodes("//table[@class=\"table_wrapper_submit\"]/tr");
                    string djzbh = listTrNodeList[0].SelectNodes("./td/input")[1].GetAttributeValue("value", "");
                    if (!CommonUtil.IsNullOrBlank(djzbh))
                    {
                        Dictionary<string, object> reportInfo = new Dictionary<string, object>();

                        reportInfo.Add("行政区划", listTrNodeList[0].SelectNodes("./td/input")[0].GetAttributeValue("value", ""));
                        reportInfo.Add("登记证编号", listTrNodeList[0].SelectNodes("./td/input")[1].GetAttributeValue("value", ""));
                        reportInfo.Add("登记表编号", listTrNodeList[1].SelectNodes("./td/input")[0].GetAttributeValue("value", ""));
                        reportInfo.Add("档案盒编号", listTrNodeList[1].SelectNodes("./td/input")[1].GetAttributeValue("value", ""));
                        reportInfo.Add("车牌号码", listTrNodeList[2].SelectNodes("./td/input")[0].GetAttributeValue("value", ""));
                        reportInfo.Add("发动机号", listTrNodeList[2].SelectNodes("./td/input")[1].GetAttributeValue("value", ""));
                        reportInfo.Add("车辆识别代码", listTrNodeList[3].SelectNodes("./td/input")[0].GetAttributeValue("value", ""));
                        reportInfo.Add("气瓶种类", listTrNodeList[3].SelectNodes("./td/input")[1].GetAttributeValue("value", ""));
                        reportInfo.Add("使用性质", listTrNodeList[4].SelectNodes("./td/input")[0].GetAttributeValue("value", ""));
                        reportInfo.Add("充装介质", listTrNodeList[4].SelectNodes("./td/input")[1].GetAttributeValue("value", ""));
                        reportInfo.Add("使用单位", listTrNodeList[5].SelectNodes("./td/input")[0].GetAttributeValue("value", ""));
                        reportInfo.Add("联系人", listTrNodeList[5].SelectNodes("./td/input")[1].GetAttributeValue("value", ""));
                        reportInfo.Add("联系电话", listTrNodeList[6].SelectNodes("./td/input")[0].GetAttributeValue("value", ""));
                        reportInfo.Add("起用日期", listTrNodeList[6].SelectNodes("./td/input")[1].GetAttributeValue("value", ""));
                        reportInfo.Add("安装单位", listTrNodeList[7].SelectNodes("./td/input")[0].GetAttributeValue("value", ""));
                        reportInfo.Add("安装日期", listTrNodeList[7].SelectNodes("./td/input")[1].GetAttributeValue("value", ""));
                        reportInfo.Add("发证日期", listTrNodeList[8].SelectNodes("./td/input")[0].GetAttributeValue("value", ""));
                        reportInfo.Add("发证机关", listTrNodeList[8].SelectNodes("./td/input")[1].GetAttributeValue("value", ""));
                        clxxEW.AddRow(reportInfo);
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

        /// <summary>
        /// 生成车辆气瓶列表
        /// </summary>
        /// <param name="listSheet"></param>
        /// <returns></returns>
        private bool GenerateCLQPLB(IListSheet listSheet)
        {
            bool succeed = true;
            string exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> clxxColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                        "登记证编号",
                        "气瓶编号",
                        "电子标签号", 
                        "制造单位", 
                        "制造日期", 
                        "审验日期", 
                        "下次检验日期"});
            string dagljlPath = Path.Combine(exportDir, this.RunPage.Project.Name + "_车辆气瓶列表.xlsx");
            Dictionary<string, string> columnFormats = new Dictionary<string, string>();
            ExcelWriter clxxEW = new ExcelWriter(dagljlPath, "List", clxxColumnDic, columnFormats);


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

                    HtmlNodeCollection qpListTrNodeList = htmlDoc.DocumentNode.SelectNodes("//table[@class=\"GridView\"]/tr");

                    for (int j = 1; j < qpListTrNodeList.Count; j++)
                    {
                        HtmlNode qpItemNode = qpListTrNodeList[j];
                        HtmlNodeCollection qpPropertyNodes = qpItemNode.SelectNodes("./td");
                        try
                        {
                            Dictionary<string, object> reportInfo = new Dictionary<string, object>();
                            reportInfo.Add("登记证编号", qpPropertyNodes[0].InnerText.Trim());
                            reportInfo.Add("气瓶编号", qpPropertyNodes[1].InnerText.Trim());
                            string dzbqh = qpPropertyNodes[2].InnerText.Trim();
                            dzbqh = CommonUtil.HtmlDecode(dzbqh).Trim();
                            reportInfo.Add("电子标签号", dzbqh);
                            reportInfo.Add("制造单位", qpPropertyNodes[3].InnerText.Trim());
                            reportInfo.Add("制造日期", qpPropertyNodes[4].InnerText.Trim());
                            reportInfo.Add("审验日期", qpPropertyNodes[5].InnerText.Trim());
                            reportInfo.Add("下次检验日期", qpPropertyNodes[6].InnerText.Trim()); 
                            clxxEW.AddRow(reportInfo);
                        }
                        catch (Exception ex)
                        {
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