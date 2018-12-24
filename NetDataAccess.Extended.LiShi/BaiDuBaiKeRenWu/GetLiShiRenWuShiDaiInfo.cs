using HtmlAgilityPack;
using NetDataAccess.Base.Common;
using NetDataAccess.Base.Config;
using NetDataAccess.Base.DB;
using NetDataAccess.Base.UI;
using NetDataAccess.Base.Writer;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace NetDataAccess.Extended.LiShi.BaiDuBaiKeRenWu
{
    public class GetLiShiRenWuShiDaiInfo : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            this.GetShiDaiInfos(listSheet);
            return true;
        }

        private void GetShiDaiInfos(IListSheet listSheet)
        {
            try
            {
                ExcelWriter renWuInfoExcelWriter = this.CreatePropertyVaueWriter();
                for (int i = 0; i < listSheet.RowCount; i++)
                {
                    Dictionary<string, string> listRow = listSheet.GetRow(i);
                    bool giveUp = "Y".Equals(listRow[SysConfig.GiveUpGrabFieldName]);
                    string pageUrl = listRow[SysConfig.DetailPageUrlFieldName];
                    string name = listRow["name"];
                    if (!giveUp)
                    {
                        try
                        {
                            HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);

                            HtmlNode itemBaseInfoNode = htmlDoc.DocumentNode.SelectSingleNode("//div[@class=\"lemmaWgt-promotion-rightPreciseAd\"]");
                            string itemId = itemBaseInfoNode.GetAttributeValue("data-lemmaid", "");
                            string itemTitle = itemBaseInfoNode.GetAttributeValue("data-lemmatitle", "");

                            HtmlNodeCollection summaryInfoNodes = htmlDoc.DocumentNode.SelectNodes("//div[@class=\"lemma-summary\"]/div[@class=\"para\"]");
                            StringBuilder summaryStrBuilder = new StringBuilder();
                            if (summaryInfoNodes != null)
                            {
                                foreach (HtmlNode summaryInfoNode in summaryInfoNodes)
                                {
                                    summaryStrBuilder.AppendLine(CommonUtil.HtmlDecode(summaryInfoNode.InnerText).Trim());
                                }
                            }

                            Dictionary<string, string> row = new Dictionary<string, string>();
                            row.Add("url", pageUrl);
                            row.Add("itemId", itemId);
                            row.Add("itemName", itemTitle);
                            row.Add("summaryInfo", summaryStrBuilder.ToString().Trim());

                            HtmlNodeCollection dtNodes = htmlDoc.DocumentNode.SelectNodes("//div[@class=\"basic-info cmn-clearfix\"]/dl/dt");
                            if (dtNodes != null)
                            {
                                List<string> oneIChaoDaiProperties = new List<string>();
                                foreach (HtmlNode dtNode in dtNodes)
                                {
                                    string pKey = CommonUtil.HtmlDecode(dtNode.InnerText).Trim().Replace(" ", "").Replace(" ", "").Replace("　", "");


                                    if (this.ShiDaiPropertyDic.ContainsKey(pKey))
                                    {
                                        string pValue = this.GetNextDDNodeText(dtNode);
                                        row[pKey] = (row.ContainsKey(pKey) ? (row[pKey] + "; ") : "") + pValue;
                                    } 

                                }
                            }

                            renWuInfoExcelWriter.AddRow(row);
                        }
                        catch (Exception ex)
                        {
                            throw ex;
                        }
                    }
                }
                renWuInfoExcelWriter.SaveToDisk(); 

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        private string GetNextDDNodeText(HtmlNode dtNode)
        {
            HtmlNode ddNode = dtNode.NextSibling;
            while (ddNode.Name.ToLower() != "dd")
            {
                ddNode = ddNode.NextSibling;
            }
            string pValue = CommonUtil.HtmlDecode(ddNode.InnerText).Trim();
            return pValue;
        }
        private List<string> _ShiDaiPropertyList=null; 
        private List<string> ShiDaiPropertyList 
        {
            get
            {
                if (this._ShiDaiPropertyList == null)
                {
                    this._ShiDaiPropertyList = new List<string>() { "所处时代", "日期", "时间", "时期", "时代", "年代", "国家", "国籍", "朝代", "出生日期", "出生时间", "去世日期", "去世时间", "逝世日期", "逝世时间" };
                }
                return this._ShiDaiPropertyList;
            }
        }
        private Dictionary<string, string> _ShiDaiPropertyDic = null;
        private Dictionary<string, string> ShiDaiPropertyDic
        {
            get
            {
                if (this._ShiDaiPropertyDic == null)
                {
                    this._ShiDaiPropertyDic = new Dictionary<string, string>();
                    for (int i = 0; i < this.ShiDaiPropertyList.Count; i++)
                    {
                        this._ShiDaiPropertyDic.Add(this.ShiDaiPropertyList[i], "");
                    }
                }
                return this._ShiDaiPropertyDic;
            }
        }

        private ExcelWriter CreatePropertyVaueWriter()
        {
            String exportDir = this.RunPage.GetExportDir();
            string resultFilePath = Path.Combine(exportDir, "百度百科_历史人物_时代相关信息.xlsx");

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("url", 0);
            resultColumnDic.Add("itemId", 1);
            resultColumnDic.Add("itemName", 2);
            resultColumnDic.Add("summaryInfo", 3);
            for (int i = 0; i < this.ShiDaiPropertyList.Count; i++)
            {
                resultColumnDic.Add(this.ShiDaiPropertyList[i], 4 + i);
            }
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }
    }
}
