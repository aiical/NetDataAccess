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
            //this.GetRelatedItemPageUrls(listSheet);
            //this.GetItemTags(listSheet);
            //this.GetItemProperties(listSheet);
            this.GetItemPropertiesAndTags(listSheet);
            return true;
        }

        public override void CheckRequestCompleteFile(string webPageText, Dictionary<string, string> listRow)
        {
            string itemName = listRow["itemName"];
            string itemUrl = listRow[SysConfig.DetailPageUrlFieldName];
            if (webPageText.Contains("您所访问的页面不存在"))
            {
                throw new Exception("获取页面失败, itemUrl = " + itemUrl + ",  itemName = " + itemName);
            }
            /*
            if (!webPageText.Contains(itemName))
            {
                throw new Exception("获取页面失败, itemUrl = " + itemUrl +",  itemName = " +itemName);
            }*/
        }

        private void AddToMoreExcelWriter(ref int pageIndex, ref ExcelWriter moreItemEW, Dictionary<string, string> moreItemRow, ref int rowCount)
        {
            rowCount++;
            if (moreItemEW == null || rowCount >= 50000)
            {
                if (moreItemEW != null)
                {
                    moreItemEW.SaveToDisk();
                }
                pageIndex++;
                moreItemEW = this.CreateMoreItemWriter(pageIndex);
                rowCount = 0;
            }
            moreItemEW.AddRow(moreItemRow);
        }

        private void GetItemTags(IListSheet listSheet)
        {
            string sourceDir = this.RunPage.GetDetailSourceFileDir();

            int rowCount = 0;
            ExcelWriter tagEW = null;

            Dictionary<string, bool> itemMaps = new Dictionary<string, bool>();
            int fileIndex = 0;

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> listRow = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(listRow[SysConfig.GiveUpGrabFieldName]);
                string fromItemUrl = listRow[SysConfig.DetailPageUrlFieldName];
                if (!giveUp)
                {
                    try
                    {
                        string localFilePath = this.RunPage.GetFilePath(fromItemUrl, sourceDir);
                        string html = FileHelper.GetTextFromFile(localFilePath, Encoding.UTF8);
                        if (!html.Contains("您所访问的页面不存在"))
                        {
                            HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                            htmlDoc.LoadHtml(html);
                            HtmlNode titleNode = htmlDoc.DocumentNode.SelectSingleNode("//dd[@class=\"lemmaWgt-lemmaTitle-title\"]/h1");
                            if (titleNode == null)
                            {
                                this.RunPage.InvokeAppendLogText("此词条不存在,百度没有收录, pageUrl = " + fromItemUrl, LogLevelType.Error, true);
                            }
                            else
                            {
                                string fromItemName = CommonUtil.HtmlDecode(titleNode.InnerText).Trim();

                                HtmlNode itemBaseInfoNode = htmlDoc.DocumentNode.SelectSingleNode("//div[@class=\"lemmaWgt-promotion-rightPreciseAd\"]");
                                string fromItemId = itemBaseInfoNode.GetAttributeValue("data-lemmaid", "");
                                string fromItemTitle = itemBaseInfoNode.GetAttributeValue("data-lemmatitle", "");

                                if (!itemMaps.ContainsKey(fromItemId))
                                {
                                    itemMaps.Add(fromItemId, true);
                                    HtmlNodeCollection tagNodes = htmlDoc.DocumentNode.SelectNodes("//dd[@id=\"open-tag-item\"]/span[@class=\"taglist\"]");
                                    List<string> tagList = new List<string>();
                                    if (tagNodes != null)
                                    {
                                        foreach (HtmlNode tagNode in tagNodes)
                                        {
                                            tagList.Add(CommonUtil.HtmlDecode(tagNode.InnerText).Trim());
                                        }
                                    }

                                    rowCount++;
                                    if (rowCount > 500000 || tagEW == null)
                                    {
                                        if (tagEW != null)
                                        {
                                            tagEW.SaveToDisk();
                                        }
                                        fileIndex++;
                                        tagEW = this.CreateItemTagsWriter(fileIndex);
                                    }

                                    Dictionary<string, string> newRow = new Dictionary<string, string>();
                                    newRow.Add("url", fromItemUrl);
                                    newRow.Add("itemId", fromItemId);
                                    newRow.Add("itemName", fromItemTitle);
                                    newRow.Add("tags", CommonUtil.StringArrayToString(tagList.ToArray(), ","));
                                    tagEW.AddRow(newRow);
                                }
                            }

                        }
                        else
                        {
                            this.RunPage.InvokeAppendLogText("放弃解析此页, 所访问的页面不存在, pageUrl = " + fromItemUrl, LogLevelType.Error, true);
                        }
                    }
                    catch (Exception ex)
                    {
                        this.RunPage.InvokeAppendLogText(ex.Message + ". 解析出错， pageUrl = " + fromItemUrl, LogLevelType.Error, true);
                        throw ex;
                    }
                }
            }

            tagEW.SaveToDisk();
        }

        private void GetItemProperties(IListSheet listSheet)
        {
            string sourceDir = this.RunPage.GetDetailSourceFileDir();

            int rowCount = 0;
            ExcelWriter propertyEW = null;

            Dictionary<string, bool> itemMaps = new Dictionary<string, bool>();
            int fileIndex = 0;
            int succeedCount = 0;

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> listRow = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(listRow[SysConfig.GiveUpGrabFieldName]);
                string fromItemUrl = listRow[SysConfig.DetailPageUrlFieldName];

                if (!giveUp)
                {
                    try
                    {
                        string localFilePath = this.RunPage.GetFilePath(fromItemUrl, sourceDir);
                        string html = FileHelper.GetTextFromFile(localFilePath, Encoding.UTF8);
                        if (!html.Contains("您所访问的页面不存在"))
                        {
                            HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                            htmlDoc.LoadHtml(html);
                            HtmlNode titleNode = htmlDoc.DocumentNode.SelectSingleNode("//dd[@class=\"lemmaWgt-lemmaTitle-title\"]/h1");
                            if (titleNode == null)
                            {
                                this.RunPage.InvokeAppendLogText("此词条不存在,百度没有收录, pageUrl = " + fromItemUrl, LogLevelType.Error, true);
                            }
                            else
                            {
                                string fromItemName = CommonUtil.HtmlDecode(titleNode.InnerText).Trim();

                                HtmlNode itemBaseInfoNode = htmlDoc.DocumentNode.SelectSingleNode("//div[@class=\"lemmaWgt-promotion-rightPreciseAd\"]");
                                string fromItemId = itemBaseInfoNode.GetAttributeValue("data-lemmaid", "");
                                string fromItemTitle = itemBaseInfoNode.GetAttributeValue("data-lemmatitle", "");

                                if (!itemMaps.ContainsKey(fromItemId))
                                {
                                    itemMaps.Add(fromItemId, true);
                                    HtmlNodeCollection dtNodes = htmlDoc.DocumentNode.SelectNodes("//div[@class=\"basic-info cmn-clearfix\"]/dl/dt");
                                    if (dtNodes != null)
                                    {
                                        List<string> itemProperties = new List<string>();
                                        foreach (HtmlNode dtNode in dtNodes)
                                        {
                                            string pKey = CommonUtil.HtmlDecode(dtNode.InnerText).Trim().Replace(" ", "").Replace(" ", "").Replace("　", "");
                                            if (!itemProperties.Contains(pKey))
                                            {
                                                itemProperties.Add(pKey);
                                            }
                                        }

                                        succeedCount++;
                                        rowCount++;
                                        if (rowCount > 500000 || propertyEW == null)
                                        {
                                            if (propertyEW != null)
                                            {
                                                propertyEW.SaveToDisk();
                                            }
                                            fileIndex++;
                                            propertyEW = this.CreateItemPropertiesWriter(fileIndex);
                                        }

                                        Dictionary<string, string> newRow = new Dictionary<string, string>();
                                        newRow.Add("url", fromItemUrl);
                                        newRow.Add("itemId", fromItemId);
                                        newRow.Add("itemName", fromItemTitle);
                                        newRow.Add("properties", CommonUtil.StringArrayToString(itemProperties.ToArray(), ","));
                                        propertyEW.AddRow(newRow);
                                    }
                                }
                                else
                                {
                                    this.RunPage.InvokeAppendLogText("(成功解析：" + succeedCount.ToString() + "个，正在处理第" + (i + 1).ToString() + "个)放弃解析此页, 已经解析过此页, pageUrl = " + fromItemUrl, LogLevelType.Error, true);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        this.RunPage.InvokeAppendLogText(ex.Message + ". 解析出错， pageUrl = " + fromItemUrl, LogLevelType.Error, true);
                        throw ex;
                    }
                }
            }

            propertyEW.SaveToDisk();
        }

        private void GetItemPropertiesAndTags(IListSheet listSheet)
        {
            string sourceDir = this.RunPage.GetDetailSourceFileDir();

            int rowCount = 0;
            ExcelWriter ptEW = null;

            Dictionary<string, bool> itemMaps = new Dictionary<string, bool>();
            int fileIndex = 0;
            int succeedCount = 0;

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> listRow = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(listRow[SysConfig.GiveUpGrabFieldName]);
                string fromItemUrl = listRow[SysConfig.DetailPageUrlFieldName];

                if (!giveUp)
                {
                    try
                    {
                        string localFilePath = this.RunPage.GetFilePath(fromItemUrl, sourceDir);
                        string html = FileHelper.GetTextFromFile(localFilePath, Encoding.UTF8);
                        if (!html.Contains("您所访问的页面不存在"))
                        {
                            HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                            htmlDoc.LoadHtml(html);
                            HtmlNode titleNode = htmlDoc.DocumentNode.SelectSingleNode("//dd[@class=\"lemmaWgt-lemmaTitle-title\"]/h1");
                            if (titleNode == null)
                            {
                                this.RunPage.InvokeAppendLogText("此词条不存在,百度没有收录, pageUrl = " + fromItemUrl, LogLevelType.Error, true);
                            }
                            else
                            {
                                string fromItemName = CommonUtil.HtmlDecode(titleNode.InnerText).Trim();

                                HtmlNode itemBaseInfoNode = htmlDoc.DocumentNode.SelectSingleNode("//div[@class=\"lemmaWgt-promotion-rightPreciseAd\"]");
                                string fromItemId = itemBaseInfoNode.GetAttributeValue("data-lemmaid", "");
                                string fromItemTitle = itemBaseInfoNode.GetAttributeValue("data-lemmatitle", "");

                                if (!itemMaps.ContainsKey(fromItemId))
                                {
                                    itemMaps.Add(fromItemId, true);
                                    HtmlNodeCollection dtNodes = htmlDoc.DocumentNode.SelectNodes("//div[@class=\"basic-info cmn-clearfix\"]/dl/dt");
                                    List<string> itemPT = new List<string>();
                                    if (dtNodes != null)
                                    {
                                        foreach (HtmlNode dtNode in dtNodes)
                                        {
                                            string pKey = CommonUtil.HtmlDecode(dtNode.InnerText).Trim().Replace(" ", "").Replace(" ", "").Replace("　", "");
                                            if (!itemPT.Contains(pKey))
                                            {
                                                itemPT.Add(pKey);
                                            }
                                        }
                                    }

                                    HtmlNodeCollection tagNodes = htmlDoc.DocumentNode.SelectNodes("//dd[@id=\"open-tag-item\"]/span[@class=\"taglist\"]");
                                    if (tagNodes != null)
                                    {
                                        foreach (HtmlNode tagNode in tagNodes)
                                        {
                                            itemPT.Add(CommonUtil.HtmlDecode(tagNode.InnerText).Trim());
                                        }
                                    }
                                    if (itemPT.Count > 0)
                                    {
                                        succeedCount++;
                                        rowCount++;
                                        if (rowCount > 500000 || ptEW == null)
                                        {
                                            if (ptEW != null)
                                            {
                                                ptEW.SaveToDisk();
                                            }
                                            fileIndex++;
                                            ptEW = this.CreateItemPTWriter(fileIndex);
                                        }

                                        Dictionary<string, string> newRow = new Dictionary<string, string>();
                                        newRow.Add("url", fromItemUrl);
                                        newRow.Add("itemId", fromItemId);
                                        newRow.Add("itemName", fromItemTitle);
                                        newRow.Add("pts", CommonUtil.StringArrayToString(itemPT.ToArray(), ","));
                                        ptEW.AddRow(newRow);
                                    }
                                }
                                else
                                {
                                    this.RunPage.InvokeAppendLogText("(成功解析：" + succeedCount.ToString() + "个，正在处理第" + (i + 1).ToString() + "个)放弃解析此页, 已经解析过此页, pageUrl = " + fromItemUrl, LogLevelType.Error, true);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        this.RunPage.InvokeAppendLogText(ex.Message + ". 解析出错， pageUrl = " + fromItemUrl, LogLevelType.Error, true);
                        throw ex;
                    }
                }
            }

            ptEW.SaveToDisk();
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

        private void GetRelatedItemPageUrls(IListSheet listSheet)
        {
            string sourceDir = this.RunPage.GetDetailSourceFileDir();

            int rowCount = 0;
            ExcelWriter moreItemEW = null;

            Dictionary<string, bool> itemMaps = new Dictionary<string, bool>();
            int pageIndex = 0;

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
                this.AddToMoreExcelWriter(ref pageIndex, ref moreItemEW, moreItemRow, ref rowCount);
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
                        string localFilePath = this.RunPage.GetFilePath(fromItemUrl, sourceDir);
                        string html = FileHelper.GetTextFromFile(localFilePath, Encoding.UTF8);
                        if (!html.Contains("您所访问的页面不存在"))
                        {
                            HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                            htmlDoc.LoadHtml(html);
                            HtmlNode titleNode = htmlDoc.DocumentNode.SelectSingleNode("//dd[@class=\"lemmaWgt-lemmaTitle-title\"]/h1");
                            if (titleNode == null)
                            {
                                this.RunPage.InvokeAppendLogText("此词条不存在,百度没有收录, pageUrl = " + fromItemUrl, LogLevelType.Error, true);
                            }
                            else
                            {
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

                                        this.AddToMoreExcelWriter(ref pageIndex, ref moreItemEW, moreItemRow, ref rowCount);
                                    }
                                }
                                this.GenerateRelatedItemFile(fromItemId, fromItemName, fromItemTitle, fromItemUrl, aNodes);
                            }

                        }
                        else
                        {
                            this.RunPage.InvokeAppendLogText("放弃解析此页, 所访问的页面不存在, pageUrl = " + fromItemUrl, LogLevelType.Error, true);
                        }
                    }
                    catch (Exception ex)
                    {
                        this.RunPage.InvokeAppendLogText(ex.Message + ". 解析出错， pageUrl = " + fromItemUrl, LogLevelType.Error, true);
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
            string resultFilePath = Path.Combine(exportDir, partDir + "/百度百科_词条关联_" + CommonUtil.ProcessFileName(fromItemTitle, "_") + "_" + fromItemId + ".xlsx");
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

        private ExcelWriter CreateMoreItemWriter(int pageIndex)
        {
            String exportDir = this.RunPage.GetExportDir();
            string resultFilePath = Path.Combine(exportDir, "百度百科_词条_详情页_" + pageIndex.ToString() + ".xlsx");

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
        private ExcelWriter CreateItemTagsWriter(int fileIndex)
        {
            String exportDir = this.RunPage.GetExportDir();
            string resultFilePath = Path.Combine(exportDir, "百度百科_词条_Tags_" + fileIndex.ToString() + ".xlsx");

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("url", 0);
            resultColumnDic.Add("itemId", 1);
            resultColumnDic.Add("itemName", 2);
            resultColumnDic.Add("tags", 3);
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }
        private ExcelWriter CreateItemPTWriter(int fileIndex)
        {
            String exportDir = this.RunPage.GetExportDir();
            string resultFilePath = Path.Combine(exportDir, "百度百科_词条_PTs_" + fileIndex.ToString() + ".xlsx");

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("url", 0);
            resultColumnDic.Add("itemId", 1);
            resultColumnDic.Add("itemName", 2);
            resultColumnDic.Add("pts", 3);
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }
        private ExcelWriter CreateItemPropertiesWriter(int fileIndex)
        {
            String exportDir = this.RunPage.GetExportDir();
            string resultFilePath = Path.Combine(exportDir, "百度百科_词条_Properties_" + fileIndex.ToString() + ".xlsx");

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("url", 0);
            resultColumnDic.Add("itemId", 1);
            resultColumnDic.Add("itemName", 2);
            resultColumnDic.Add("properties", 3);
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }
    }
}