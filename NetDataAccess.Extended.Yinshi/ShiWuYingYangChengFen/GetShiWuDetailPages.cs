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

namespace NetDataAccess.Extended.Yinshi.ShiWuYingYangChengFen
{
    public class GetShiWuDetailPages : ExternalRunWebPage
    { 
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            try
            {
                this.GetList(listSheet); 
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return true;
        }

        private void GetList(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>(); 
            resultColumnDic.Add("所属类别", 0);
            resultColumnDic.Add("食物名称", 1);
            resultColumnDic.Add("营养成分", 2);
            resultColumnDic.Add("含量", 3);
            resultColumnDic.Add("单位", 4);
            string resultFilePath = Path.Combine(exportDir, "食物营养成分.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null); 
            
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"];
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                { 
                    try
                    {
                        HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                        HtmlNodeCollection itemTitleNodes = htmlDoc.DocumentNode.SelectNodes("//table[@class=\"yingyang wkbx\"]/tr/th");
                        HtmlNodeCollection itemValueNodes = htmlDoc.DocumentNode.SelectNodes("//table[@class=\"yingyang wkbx\"]/tr/td");

                        if (itemTitleNodes.Count != itemValueNodes.Count)
                        {
                            throw new Exception("name = " + row["名称"] + ", detailUrl = " + detailUrl + "成分含量信息格式错误");
                        }

                        for (int j = 0; j < itemTitleNodes.Count; j++)
                        {
                            HtmlNode itemTitleNode = itemTitleNodes[j];
                            HtmlNode itemValueNode = itemValueNodes[j];
                            string yycf = itemTitleNode.InnerText.Trim();
                            string hl = "";
                            string dw = "";
                            foreach (HtmlNode valueNode in itemValueNode.ChildNodes)
                            {
                                if (valueNode.NodeType == HtmlNodeType.Text)
                                {
                                    hl = valueNode.InnerText.Trim();
                                }
                                else if(valueNode.NodeType == HtmlNodeType.Element)
                                {
                                    if (valueNode.Name.ToLower() == "span")
                                    {
                                        dw = valueNode.InnerText.Replace("(", "").Replace(")", "").Trim();
                                    }
                                }
                            }

                            Dictionary<string, string> f2vs = new Dictionary<string, string>();
                            f2vs.Add("所属类别", row["所属类别"]);
                            f2vs.Add("食物名称", row["名称"]);
                            f2vs.Add("营养成分", yycf);
                            f2vs.Add("含量", hl);
                            f2vs.Add("单位", dw);
                            resultEW.AddRow(f2vs);
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