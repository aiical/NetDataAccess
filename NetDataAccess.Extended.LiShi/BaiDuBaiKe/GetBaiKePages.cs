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

namespace NetDataAccess.Extended.LiShi.BaiDuBaiKe
{
    public class GetBaiKePages : ExternalRunWebPage
    { 
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            this.GetRelatedItemPageUrls(listSheet); 
            return true;
        }

        public override void CheckRequestCompleteFile(string webPageText, Dictionary<string, string> listRow)
        {
            string itemName = listRow["itemName"];
            string itemUrl = listRow[SysConfig.DetailPageUrlFieldName];
            if (!webPageText.Contains(itemName))
            {
                throw new Exception("获取页面失败, itemUrl = " + itemUrl +",  itemName = " +itemName);
            }
        }

        private void GetRelatedItemPageUrls(IListSheet listSheet)
        {
            ExcelWriter moreItemEW = this.CreateMoreItemWriter();

            Dictionary<string, bool> itemMaps = new Dictionary<string, bool>();

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> listRow = listSheet.GetRow(i);
                string fromItemUrl = listRow[SysConfig.DetailPageUrlFieldName];

                itemMaps.Add(fromItemUrl, true);

                Dictionary<string, string> moreItemRow = new Dictionary<string, string>();
                moreItemRow.Add("detailPageUrl", fromItemUrl);
                moreItemRow.Add("detailPageName", fromItemUrl);
                moreItemRow.Add("itemId", listRow["itemId"]);
                moreItemRow.Add("itemName", listRow["itemName"]);
                moreItemEW.AddRow(moreItemRow);
            }

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> listRow = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(listRow[SysConfig.GiveUpGrabFieldName]);
                string fromItemUrl = listRow[SysConfig.DetailPageUrlFieldName];
                if (!giveUp)
                {
                    try
                    {
                        HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);

                        HtmlNode titleNode = htmlDoc.DocumentNode.SelectSingleNode("//dd[@class=\"lemmaWgt-lemmaTitle-title\"]/h1");
                        string fromItemName = CommonUtil.HtmlDecode(titleNode.InnerText).Trim();

                        HtmlNode itemBaseInfoNode = htmlDoc.DocumentNode.SelectSingleNode("//div[@class=\"lemmaWgt-promotion-rightPreciseAd\"]");
                        string fromItemId = itemBaseInfoNode.GetAttributeValue("data-lemmaid", "");
                        string fromItemTitle = itemBaseInfoNode.GetAttributeValue("data-lemmatitle", "");

                        HtmlNodeCollection aNodes = htmlDoc.DocumentNode.SelectNodes("//a");
                        for (int j = 0; j < aNodes.Count; j++)
                        {
                            HtmlNode aNode = aNodes[j];
                            string toItemUrl = aNode.GetAttributeValue("href", "");
                            string toItemId = aNode.GetAttributeValue("data-lemmaid", "");
                            string toItemName = CommonUtil.HtmlDecode(aNode.InnerText).Trim();
                            string toItemFullUrl = "https://baike.baidu.com" + toItemUrl;
                            if (toItemUrl.StartsWith("/item/") && !itemMaps.ContainsKey(toItemFullUrl) && this.IsInMainContent(aNode))
                            {
                                itemMaps.Add(toItemFullUrl, true);

                                Dictionary<string, string> moreItemRow = new Dictionary<string, string>();
                                moreItemRow.Add("detailPageUrl", toItemFullUrl);
                                moreItemRow.Add("detailPageName", toItemFullUrl);
                                moreItemRow.Add("itemId", toItemId);
                                moreItemRow.Add("itemName", toItemName);

                                moreItemEW.AddRow(moreItemRow);
                            }
                        }

                        this.GenerateRelatedItemFile(fromItemId, fromItemName, fromItemTitle, fromItemUrl, aNodes);
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
            }

            moreItemEW.SaveToDisk();
        }

        private bool IsInMainContent(HtmlNode aNode)
        {
            HtmlNode parentNode = aNode.ParentNode;
            while (parentNode != null)
            {
                if (parentNode.GetAttributeValue("class", "") == "main-content")
                {
                    return true;
                }
                parentNode = parentNode.ParentNode;
            }
            return false;
        }

        private void GenerateRelatedItemFile(string fromItemId, string fromItemName, string fromItemTitle, string fromItemUrl, HtmlNodeCollection aNodes)
        { 
            String exportDir = this.RunPage.GetExportDir();
            string partDir = CommonUtil.MD5Crypto(fromItemTitle + "_" + fromItemId).Substring(0, 4);
            string resultFilePath = Path.Combine(exportDir, partDir + "/百度百科_词条关联_" + fromItemTitle + "_" + fromItemId + ".xlsx");
            if (!File.Exists(resultFilePath))
            {
                Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
                resultColumnDic.Add("fromItemUrl", 0);
                resultColumnDic.Add("fromItemId", 1);
                resultColumnDic.Add("fromItemName", 2);
                resultColumnDic.Add("fromItemTitle", 3);
                resultColumnDic.Add("toItemUrl", 4);
                resultColumnDic.Add("toItemId", 5);
                resultColumnDic.Add("toItemName", 6);
                ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
                Dictionary<string, bool> itemMaps = new Dictionary<string, bool>();
                 
                for (int j = 0; j < aNodes.Count; j++)
                {
                    HtmlNode aNode = aNodes[j];
                    string toItemUrl = aNode.GetAttributeValue("href", "");
                    string toItemId = aNode.GetAttributeValue("data-lemmaid", "");
                    string toItemName = CommonUtil.HtmlDecode(aNode.InnerText).Trim();
                    if (toItemUrl.StartsWith("/item/") && !itemMaps.ContainsKey(toItemUrl) && this.IsInMainContent(aNode))
                    {
                        itemMaps.Add(toItemUrl, true);

                        string toItemFullUrl = "https://baike.baidu.com" + toItemUrl;
                        Dictionary<string, string> relatedItemRow = new Dictionary<string, string>();
                        relatedItemRow.Add("fromItemUrl", fromItemUrl);
                        relatedItemRow.Add("fromItemId", fromItemId);
                        relatedItemRow.Add("fromItemName", fromItemName);
                        relatedItemRow.Add("fromItemTitle", fromItemTitle);
                        relatedItemRow.Add("toItemUrl", toItemFullUrl);
                        relatedItemRow.Add("toItemId", toItemId);
                        relatedItemRow.Add("toItemName", toItemName);

                        resultEW.AddRow(relatedItemRow);
                    }
                }
                resultEW.SaveToDisk();
            }
        }

        private ExcelWriter CreateMoreItemWriter()
        {
            String exportDir = this.RunPage.GetExportDir();
            string resultFilePath = Path.Combine(exportDir, "百度百科_词条_详情页.xlsx");

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("itemId", 5);
            resultColumnDic.Add("itemName", 6);
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }         
    }
}