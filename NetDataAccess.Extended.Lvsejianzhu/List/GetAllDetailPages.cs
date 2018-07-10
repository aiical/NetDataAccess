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

namespace NetDataAccess.Extended.Lvsejianzhu
{
    /// <summary>
    /// GetAllDetailPages
    /// </summary>
    [System.Runtime.InteropServices.ComVisibleAttribute(true)]
    public class GetAllDetailPages : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            return this.GetShopInfos(listSheet);
        }

        private string GetShopUrl(string url)
        {
            int qIndex = url.IndexOf("?");
            if (qIndex >= 0)
            {
                return url.Substring(0, qIndex).Replace("/", "");
            }
            else
            {
                return url.Replace("/", "");
            }
        }

        private string GetShopId(string shopUrl)
        {
            int dIndex = shopUrl.IndexOf(".");
            return "https://" + shopUrl.Substring(0, dIndex);
        }

        private string GetShopType(string shopUrl)
        {
            int ldIndex = shopUrl.LastIndexOf(".");
            string tempStr = shopUrl.Substring(0, ldIndex);
            int ddIndex = tempStr.LastIndexOf(".");
            if (ddIndex >= 0)
            {
                return tempStr.Substring(ddIndex + 1).ToLower();
            }
            else
            {
                return tempStr.ToLower();
            }
        }

        private string GetShopProductListPageUrl(string shopId)
        {
            string url = "https://" + shopId + ".taobao.com/search.htm";
            return url;
        }

        /// <summary>
        /// 获取列表页里的店铺信息
        /// </summary>
        /// <param name="listSheet"></param>
        /// <returns></returns>
        private bool GetShopInfos(IListSheet listSheet)
        {
            bool succeed = true;
            string exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> columnDic = CommonUtil.InitStringIndexDic(new string[]{
                    "项目名称",
                    "项目地点", 
                    "认证类别",
                    "主要功能", 
                    "投资单位",
                    "咨询单位",
                    "认证时间",
                    "建筑面积",
                    "设计单位",
                    "施工单位",
                    "项目简介"});
            string shopFirstPageUrlFilePath = Path.Combine(exportDir, "项目详情.xlsx");
            ExcelWriter ew = new ExcelWriter(shopFirstPageUrlFilePath, "List", columnDic, null);
             
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"]; 
                string pageNum = row["pageNum"]; 
                string localFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir);
                TextReader tr = null;

                try
                {
                    tr = new StreamReader(localFilePath, Encoding.GetEncoding(((Proj_Detail_SingleLine)this.RunPage.Project.DetailGrabInfoObject).Encoding));
                    string webPageHtml = tr.ReadToEnd();

                    HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                    htmlDoc.LoadHtml(webPageHtml);

                    HtmlNode mainNode = htmlDoc.DocumentNode.SelectSingleNode("//div[@id=\"waitao\"]/table[3]/tr/td[2]");

                    this.GetProjectItem(mainNode, pageNum, ew); 
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
            ew.SaveToDisk();
            return succeed;
        }

        private void GetProjectItem(HtmlNode mainNode, string pageNum, ExcelWriter ew)
        {
            Dictionary<string, object> projectInfo = new Dictionary<string, object>();
            string xmmc = CommonUtil.HtmlDecode(mainNode.SelectSingleNode("./table[1]/tr/td").FirstChild.InnerText).Trim();
            projectInfo.Add("项目名称", xmmc);

            HtmlNode baseInfoNode = mainNode.SelectSingleNode("./table[2]/tr/td/table/tr/td[2]/div[@id=\"leftinfo1\"]/table[2]/tr/td[1]");

            string xmdd = "";
            HtmlNode addressNode = baseInfoNode.SelectSingleNode("./table[1]/tr/td[1]");
            if (addressNode != null)
            {
                xmdd = CommonUtil.HtmlDecode(addressNode.InnerText).Trim().Replace("项目地点：", "");
            }
            projectInfo.Add("项目地点", xmdd);

            HtmlNodeCollection baseInnerNodes = baseInfoNode.SelectNodes("./table[2]/tr/td");
            foreach (HtmlNode baseInnerNode in baseInnerNodes)
            {
                string fieldName = "";
                foreach (HtmlNode node in baseInnerNode.ChildNodes)
                {
                    string nodeText = CommonUtil.HtmlDecode(node.InnerText).Trim();
                    string checkStr = nodeText.Length > 5 ? nodeText.Substring(0, 5) : nodeText;
                    switch (checkStr)
                    {
                        case "认证类别：":
                            fieldName = "认证类别";
                            projectInfo.Add(fieldName, "");
                            break;
                        case "主要功能：":
                            fieldName = "主要功能";
                            projectInfo.Add(fieldName, "");
                            break;
                        case "投资单位：":
                            fieldName = "投资单位";
                            projectInfo.Add(fieldName, "");
                            break;
                        case "咨询单位：":
                            fieldName = "咨询单位";
                            projectInfo.Add(fieldName, "");
                            break;
                        case "认证时间：":
                            fieldName = "认证时间";
                            projectInfo.Add(fieldName, "");
                            break;
                        case "建筑面积：":
                            fieldName = "建筑面积";
                            projectInfo.Add(fieldName, nodeText.Substring(6));
                            
                            break;
                        case "设计单位：":
                            fieldName = "设计单位";
                            projectInfo.Add(fieldName, "");
                            break;
                        case "施工单位：":
                            fieldName = "施工单位";
                            projectInfo.Add(fieldName, "");
                            break;
                        default:
                            if (fieldName != "")
                            {
                                projectInfo[fieldName] = projectInfo[fieldName] + nodeText;
                            }
                            break;
                    }
                }
            }
            HtmlNode xmjjInfoNode = mainNode.SelectSingleNode("./table[2]/tr/td/table/tr/td[2]/div[@id=\"leftinfo2\"]/div[@id=\"xiazai1\"]");
            string xmjj = CommonUtil.HtmlDecode(xmjjInfoNode.InnerText).Trim();
            projectInfo.Add("项目简介", xmjj);

            ew.AddRow(projectInfo);
        } 
    }
}