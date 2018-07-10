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
    public class GetBenkeZhuanyeDetailPage_sczsxx_com : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            try
            {
                this.GetZhuanyeInfo(listSheet); 
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return true;
        }

        private void GetZhuanyeInfo(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("学位", 0);
            resultColumnDic.Add("学科分类", 1);
            resultColumnDic.Add("一级学科", 2);
            resultColumnDic.Add("专业", 3);
            resultColumnDic.Add("专业代码", 4);
            resultColumnDic.Add("专业培养目标", 5);
            resultColumnDic.Add("专业学习方向", 6);
            resultColumnDic.Add("毕业生应获得以下几方面的知识和能力", 7);
            resultColumnDic.Add("主要课程", 8);
            resultColumnDic.Add("主要实践性教学环节", 9);
            resultColumnDic.Add("修业年限", 10);
            resultColumnDic.Add("授予学位", 11);
            resultColumnDic.Add("就业方向", 12);
            resultColumnDic.Add("url", 13);
            string resultFilePath = Path.Combine(exportDir, "大学本科专业信息_sczsxx_com.xlsx");
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
                        HtmlNode mainNode = htmlDoc.DocumentNode.SelectSingleNode("//div[@class=\"content_txt\"]");

                        if (mainNode == null)
                        {
                            throw new Exception("未找到任何详情内容div.class=content_txt");

                        }
                        else
                        {

                            Dictionary<string, string> f2vs = new Dictionary<string, string>();
                            try
                            {
                                f2vs.Add("url", detailUrl);
                                f2vs.Add("学位", row["学位"]);
                                f2vs.Add("学科分类", row["学科分类"]);
                                f2vs.Add("一级学科", row["一级学科"]);
                                f2vs.Add("专业", row["专业"]);

                                HtmlNode fullNameNode = mainNode.SelectSingleNode("./h4");
                                if (fullNameNode != null && fullNameNode.InnerText.Contains("（本科）"))
                                {
                                    HtmlNodeCollection pNodes = mainNode.SelectNodes("./p");
                                    foreach (HtmlNode pNode in pNodes)
                                    {
                                        try
                                        {
                                            string text = pNode.InnerText.Trim();
                                            if (text.Contains("专业代码："))
                                            {
                                                int codeBeginIndex = text.LastIndexOf("：") + 1;
                                                string zhuanyeCode = text.Substring(codeBeginIndex).Trim();
                                                f2vs.Add("专业代码", zhuanyeCode);
                                            }

                                            if (text.StartsWith("培养目标")
                                                || text.StartsWith("专业培养目标")
                                                || text.StartsWith("业务培养目标")
                                                || text.StartsWith("业务培养要求"))
                                            {
                                                if (!f2vs.ContainsKey("专业培养目标"))
                                                {
                                                    int splitIndex = text.IndexOf("：");
                                                    splitIndex = splitIndex == -1 ? text.IndexOf(":") : splitIndex; 
                                                    string content = text.Substring(splitIndex + 1).Trim();
                                                    f2vs.Add("专业培养目标", content);
                                                }
                                            }
                                            else if (text.StartsWith("专业学习方向")
                                               || text.StartsWith("毕业生应获得以下几方面的知识和能力")
                                               || text.StartsWith("主要实践性教学环节"))
                                            {
                                                int splitIndex = text.IndexOf("：");
                                                splitIndex = splitIndex == -1 ? text.IndexOf(":") : splitIndex;
                                                string itemTitle = text.Substring(0, splitIndex).Trim();
                                                string content = text.Substring(splitIndex + 1).Trim();
                                                f2vs.Add(itemTitle, content);
                                            }
                                            else if (text.StartsWith("主要课程"))
                                            {
                                                int splitIndex = text.IndexOf("：");
                                                splitIndex = splitIndex == -1 ? text.IndexOf(":") : splitIndex;
                                                string content = text.Substring(splitIndex + 1).Trim();
                                                f2vs.Add("主要课程", content);
                                            }
                                            else if (text.StartsWith("修业年限")
                                                || text.StartsWith("学制"))
                                            {
                                                int splitIndex = text.IndexOf("：");
                                                splitIndex = splitIndex == -1 ? text.IndexOf(":") : splitIndex;
                                                string content = text.Substring(splitIndex + 1).Trim();
                                                f2vs.Add("修业年限", content);
                                            }
                                            else if (text.StartsWith("授予学位："))
                                            {
                                                int splitIndex = text.IndexOf("：");
                                                splitIndex = splitIndex == -1 ? text.IndexOf(":") : splitIndex;
                                                string itemTitle = text.Substring(0, splitIndex).Trim();
                                                string content = text.Substring(splitIndex + 1).Trim();
                                                f2vs.Add(itemTitle, content);
                                            }
                                            else if (text.StartsWith("就业方向"))
                                            {
                                                int splitIndex = text.IndexOf("：");
                                                splitIndex = splitIndex == -1 ? text.IndexOf(":") : splitIndex;
                                                string itemTitle = text.Substring(0, splitIndex).Trim();
                                                string content = text.Substring(splitIndex + 1).Trim();
                                                f2vs.Add(itemTitle, content);
                                            }
                                            else if (text.StartsWith("授予"))
                                            {
                                                string content = text.Substring(2).Trim();
                                                f2vs.Add("授予学位", content);
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            throw ex;
                                        }
                                    }
                                }

                                resultEW.AddRow(f2vs);
                            }
                            catch (Exception ex)
                            {
                                throw ex;
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