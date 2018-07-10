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

namespace NetDataAccess.Extended.Jiaoyu.Zhuanye
{
    public class GetBenkeZhuanyeListPage_sczsxx_com : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            try
            {
                this.GetZhuanyeList(listSheet); 
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return true;
        }
         
        private void GetZhuanyeList(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("学位", 5);
            resultColumnDic.Add("学科分类", 6); 
            resultColumnDic.Add("一级学科", 7);
            resultColumnDic.Add("专业", 8);
            string resultFilePath = Path.Combine(exportDir, "教育_专业_本科_详情_sczsxx_com.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null); 
             
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"];
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string localFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir);

                    try
                    {
                        HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                        HtmlNode benkeNode = htmlDoc.DocumentNode.SelectSingleNode("//a[@name=\"zybenke\"]");

                        if (benkeNode == null)
                        {
                            throw new Exception("未找到本科标签");
                        }
                        else
                        {
                            while (benkeNode.GetAttributeValue("class", "") != "listtype-5 clearfix")
                            {
                                benkeNode = benkeNode.NextSibling;
                            }
                            if (benkeNode == null)
                            {
                                throw new Exception("未找到专业列表");
                            }
                            else
                            {
                                HtmlNodeCollection daleiNodes = benkeNode.SelectNodes("./a");
                                foreach (HtmlNode daleiNode in daleiNodes)
                                {
                                    HtmlNode tempNode = daleiNode;
                                    string daleiName = "";
                                    string xiaoleiName = "";
                                    string zhuanyeName = "";
                                    string zhuanyeUrl = "";
                                    while (1 == 1)
                                    {
                                        tempNode = tempNode.NextSibling;
                                        if (tempNode != null)
                                        {
                                            string tagName = tempNode.Name.ToUpper();
                                            if (tagName == "H4")
                                            {
                                                string text = tempNode.InnerText.Trim();
                                                int endIndex = text.IndexOf("(");
                                                daleiName = text.Substring(0, endIndex);
                                            }
                                            else if (tagName == "H5")
                                            {
                                                xiaoleiName = tempNode.InnerText.Trim();
                                            }
                                            else if (tagName == "UL")
                                            {
                                                HtmlNodeCollection zhuanyeNodes = tempNode.SelectNodes("./li/a");
                                                foreach (HtmlNode zhuanyeNode in zhuanyeNodes)
                                                {
                                                    zhuanyeName = zhuanyeNode.InnerText.Trim();
                                                    zhuanyeUrl = "http://www.sczsxx.com" + zhuanyeNode.GetAttributeValue("href", "");

                                                    Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                                    f2vs.Add("detailPageUrl", zhuanyeUrl);
                                                    f2vs.Add("detailPageName", zhuanyeUrl);
                                                    f2vs.Add("学位", "本科");
                                                    f2vs.Add("学科分类", daleiName);
                                                    f2vs.Add("一级学科", xiaoleiName);
                                                    f2vs.Add("专业", zhuanyeName);
                                                    resultEW.AddRow(f2vs);
                                                }
                                            }
                                            else if (tagName == "A")
                                            {
                                                break;
                                            }
                                        }
                                        else
                                        {
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    } 
                }
            }
            resultEW.SaveToDisk();
        }
    }
}