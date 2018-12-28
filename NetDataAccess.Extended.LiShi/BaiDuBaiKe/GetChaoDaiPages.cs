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
    public class GetChaoDaiPages : ExternalRunWebPage
    { 
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            this.GetItemBaseInfos(listSheet);

            this.GetRelatedItemPageUrls(listSheet);

            this.GetChaoDaiProperties(listSheet);

            this.GetChaoDaiRemainProperties(listSheet);

            return true;
        }



        private void GetItemBaseInfos(IListSheet listSheet)
        {
            string sourceDir = this.RunPage.GetDetailSourceFileDir();

            int rowCount = 0;
            ExcelWriter baseInfoEW = null;

            Dictionary<string, bool> itemMaps = new Dictionary<string, bool>();
            int fileIndex = 0;

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> listRow = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(listRow[SysConfig.GiveUpGrabFieldName]);
                string itemUrl = listRow[SysConfig.DetailPageUrlFieldName];
                if (!giveUp)
                {
                    try
                    {
                        string localFilePath = this.RunPage.GetFilePath(itemUrl, sourceDir);
                        string html = FileHelper.GetTextFromFile(localFilePath, Encoding.UTF8);
                        if (!html.Contains("您所访问的页面不存在"))
                        {
                            HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                            htmlDoc.LoadHtml(html);
                            HtmlNode titleNode = htmlDoc.DocumentNode.SelectSingleNode("//dd[@class=\"lemmaWgt-lemmaTitle-title\"]/h1");
                            if (titleNode == null)
                            {
                                this.RunPage.InvokeAppendLogText("此词条不存在,百度没有收录, pageUrl = " + itemUrl, LogLevelType.Error, true);
                            }
                            else
                            {
                                string fromItemName = CommonUtil.HtmlDecode(titleNode.InnerText).Trim();

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


                                if (!itemMaps.ContainsKey(itemId))
                                {
                                    itemMaps.Add(itemId, true);
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
                                    if (rowCount > 500000 || baseInfoEW == null)
                                    {
                                        if (baseInfoEW != null)
                                        {
                                            baseInfoEW.SaveToDisk();
                                        }
                                        fileIndex++;
                                        baseInfoEW = this.CreateItemBaseInfoWriter(fileIndex);
                                    }

                                    Dictionary<string, string> newRow = new Dictionary<string, string>();
                                    newRow.Add("url", itemUrl);
                                    newRow.Add("itemId", itemId);
                                    newRow.Add("itemName", itemTitle);
                                    newRow.Add("summaryInfo", summaryStrBuilder.ToString().Trim()); 
                                    baseInfoEW.AddRow(newRow);
                                }
                            }

                        }
                        else
                        {
                            this.RunPage.InvokeAppendLogText("放弃解析此页, 所访问的页面不存在, pageUrl = " + itemUrl, LogLevelType.Error, true);
                        }
                    }
                    catch (Exception ex)
                    {
                        this.RunPage.InvokeAppendLogText(ex.Message + ". 解析出错， pageUrl = " + itemUrl, LogLevelType.Error, true);
                        throw ex;
                    }
                }
            }

            baseInfoEW.SaveToDisk();
        }

        private void GetChaoDaiProperties(IListSheet listSheet)
        {
            try
            {
                List<string> propertyColumnNames = new List<string>();

                ExcelWriter chaoDaiInfoExcelWriter = this.CreateChaoDaiPropertyListWriter();
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

                            HtmlNodeCollection dtNodes = htmlDoc.DocumentNode.SelectNodes("//div[@class=\"basic-info cmn-clearfix\"]/dl/dt");
                            if (dtNodes != null)
                            {
                                List<string> oneIChaoDaiProperties = new List<string>();
                                foreach (HtmlNode dtNode in dtNodes)
                                {
                                    string pKey = CommonUtil.HtmlDecode(dtNode.InnerText).Trim().Replace(" ", "").Replace(" ", "").Replace("　", "");
                                    string pValue = this.GetNextDDNodeText(dtNode);

                                    int sameNamePKeyCount = 1;
                                    string newPKey = pKey;
                                    while (oneIChaoDaiProperties.Contains(newPKey))
                                    {
                                        sameNamePKeyCount++;
                                        newPKey = pKey + "_" + sameNamePKeyCount.ToString();
                                    }
                                    oneIChaoDaiProperties.Add(newPKey);

                                    if (!propertyColumnNames.Contains(newPKey))
                                    {
                                        propertyColumnNames.Add(newPKey);
                                    }

                                    Dictionary<string, string> row = new Dictionary<string, string>();

                                    row.Add("name", name);
                                    row.Add("pKey", newPKey);
                                    row.Add("pValue", pValue);
                                    row.Add("url", pageUrl);

                                    chaoDaiInfoExcelWriter.AddRow(row);

                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            throw ex;
                        }
                    }
                }
                chaoDaiInfoExcelWriter.SaveToDisk();

                ExcelWriter chaoDaiColumnPropertyExcelWriter = this.CreateChaoDaiColumnPropertyListWriter(propertyColumnNames);
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
                            HtmlNodeCollection dtNodes = htmlDoc.DocumentNode.SelectNodes("//div[@class=\"basic-info cmn-clearfix\"]/dl/dt");
                            Dictionary<string, string> row = new Dictionary<string, string>();
                            row.Add("name", name);
                            row.Add("url", pageUrl);
                            if (dtNodes != null)
                            {
                                List<string> oneIChaoDaiProperties = new List<string>();
                                foreach (HtmlNode dtNode in dtNodes)
                                {
                                    string pKey = CommonUtil.HtmlDecode(dtNode.InnerText).Trim().Replace(" ", "").Replace(" ", "").Replace("　", "");
                                    string pValue = this.GetNextDDNodeText(dtNode);

                                    int sameNamePKeyCount = 1;
                                    string newPKey = pKey;
                                    while (oneIChaoDaiProperties.Contains(newPKey))
                                    {
                                        sameNamePKeyCount++;
                                        newPKey = pKey + "_" + sameNamePKeyCount.ToString();
                                    }
                                    oneIChaoDaiProperties.Add(newPKey);

                                    row.Add(newPKey, pValue);
                                }
                            }

                            chaoDaiColumnPropertyExcelWriter.AddRow(row);
                        }
                        catch (Exception ex)
                        {
                            throw ex;
                        }
                    }
                }
                chaoDaiColumnPropertyExcelWriter.SaveToDisk();

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        /// <summary>
        /// 保留部分属性
        /// </summary>
        /// <param name="listSheet"></param>
        private void GetChaoDaiRemainProperties(IListSheet listSheet)
        {
            try
            {
                string[] parameters = this.Parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                string columnMapFilePath = parameters[0];

                ExcelReader columnMapER = new ExcelReader(columnMapFilePath, "朝代属性");
                int rowCount = columnMapER.GetRowCount();
                Dictionary<string, string> columnAliasToColumns = new Dictionary<string, string>();
                for (int i = 0; i < rowCount; i++)
                {
                    Dictionary<string,string> columnRow = columnMapER.GetFieldValues(i);
                    string columnName = columnRow["column"].Trim();
                    columnAliasToColumns.Add(columnName, columnName);

                    string aliasColumnsStr = columnRow["aliasColumns"];
                    string[] aliasColumns = aliasColumnsStr.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (string alias in aliasColumns)
                    {
                        columnAliasToColumns.Add(alias.Trim(), columnName);
                    }
                }

                List<string> propertyColumnNames = new List<string>();

                ExcelWriter chaoDaiInfoExcelWriter = this.CreateChaoDaiRemainPropertyListWriter();
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
                            HtmlNodeCollection dtNodes = htmlDoc.DocumentNode.SelectNodes("//div[@class=\"basic-info cmn-clearfix\"]/dl/dt");
                            if (dtNodes != null)
                            {
                                List<string> oneIChaoDaiProperties = new List<string>();
                                foreach (HtmlNode dtNode in dtNodes)
                                {
                                    string pKey = CommonUtil.HtmlDecode(dtNode.InnerText).Trim().Replace(" ", "").Replace(" ", "").Replace("　", "");
                                    string pValue = this.GetNextDDNodeText(dtNode);

                                    int sameNamePKeyCount = 1;
                                    string newPKey = pKey;
                                    while (oneIChaoDaiProperties.Contains(newPKey))
                                    {
                                        sameNamePKeyCount++;
                                        newPKey = pKey + "_" + sameNamePKeyCount.ToString();
                                    }
                                    oneIChaoDaiProperties.Add(newPKey);

                                    if (!propertyColumnNames.Contains(newPKey) &&  columnAliasToColumns.ContainsValue(newPKey))
                                    {
                                        propertyColumnNames.Add(newPKey);
                                    }

                                    if (columnAliasToColumns.ContainsKey(newPKey))
                                    {
                                        string columnName = columnAliasToColumns[newPKey];

                                        Dictionary<string, string> row = new Dictionary<string, string>();

                                        row.Add("name", name);
                                        row.Add("pKey", columnName);
                                        row.Add("pValue", pValue);
                                        row.Add("url", pageUrl);

                                        chaoDaiInfoExcelWriter.AddRow(row);
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
                chaoDaiInfoExcelWriter.SaveToDisk();

                ExcelWriter chaoDaiColumnPropertyExcelWriter = this.CreateChaoDaiRemainColumnPropertyListWriter(propertyColumnNames);
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
                            HtmlNodeCollection dtNodes = htmlDoc.DocumentNode.SelectNodes("//div[@class=\"basic-info cmn-clearfix\"]/dl/dt");
                            Dictionary<string, string> row = new Dictionary<string, string>();
                            row.Add("name", name);
                            row.Add("url", pageUrl);
                            if (dtNodes != null)
                            {
                                List<string> oneIChaoDaiProperties = new List<string>();
                                foreach (HtmlNode dtNode in dtNodes)
                                {
                                    string pKey = CommonUtil.HtmlDecode(dtNode.InnerText).Trim().Replace(" ", "").Replace(" ", "").Replace("　", "");
                                    string pValue = this.GetNextDDNodeText(dtNode);

                                    int sameNamePKeyCount = 1;
                                    string newPKey = pKey;
                                    while (oneIChaoDaiProperties.Contains(newPKey))
                                    {
                                        sameNamePKeyCount++;
                                        newPKey = pKey + "_" + sameNamePKeyCount.ToString();
                                    }
                                    oneIChaoDaiProperties.Add(newPKey);

                                    if (columnAliasToColumns.ContainsKey(newPKey))
                                    {
                                        string columnName = columnAliasToColumns[newPKey];
                                        if (row.ContainsKey(columnName))
                                        {
                                            row[columnName] = row[columnName] + ";" + pValue;
                                        }
                                        else
                                        {
                                            row.Add(columnName, pValue);
                                        }
                                    }
                                }
                            }

                            chaoDaiColumnPropertyExcelWriter.AddRow(row);
                        }
                        catch (Exception ex)
                        {
                            throw ex;
                        }
                    }
                }
                chaoDaiColumnPropertyExcelWriter.SaveToDisk();

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

        private ExcelWriter CreateChaoDaiPropertyListWriter()
        {
            String exportDir = this.RunPage.GetExportDir();
            string resultFilePath = Path.Combine(exportDir, "百度百科_朝代_属性值.xlsx");

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("name", 0);
            resultColumnDic.Add("pKey", 1);
            resultColumnDic.Add("pValue", 2);
            resultColumnDic.Add("url", 3);
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }

        private ExcelWriter CreateChaoDaiColumnPropertyListWriter(List<string> propertyColumnNames)
        {
            String exportDir = this.RunPage.GetExportDir();
            string resultFilePath = Path.Combine(exportDir, "百度百科_朝代_属性值宽表.xlsx");

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("name", 0);
            resultColumnDic.Add("url", 1);
            for (int i = 0; i < propertyColumnNames.Count; i++)
            {
                resultColumnDic.Add(propertyColumnNames[i], i + 2);
            }
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }

        private ExcelWriter CreateChaoDaiRemainPropertyListWriter()
        {
            String exportDir = this.RunPage.GetExportDir();
            string resultFilePath = Path.Combine(exportDir, "百度百科_朝代_属性值_合并同义属性.xlsx");

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("name", 0);
            resultColumnDic.Add("pKey", 1);
            resultColumnDic.Add("pValue", 2);
            resultColumnDic.Add("url", 3);
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }

        private ExcelWriter CreateChaoDaiRemainColumnPropertyListWriter(List<string> propertyColumnNames)
        {
            String exportDir = this.RunPage.GetExportDir();
            string resultFilePath = Path.Combine(exportDir, "百度百科_朝代_属性值宽表_合并同义属性.xlsx");

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("name", 0);
            resultColumnDic.Add("url", 1);
            for (int i = 0; i < propertyColumnNames.Count; i++)
            {
                resultColumnDic.Add(propertyColumnNames[i], i + 2);
            }
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }         

        private void GetRelatedItemPageUrls(IListSheet listSheet)
        {
            ExcelWriter moreItemEW = this.CreateMoreItemWriter();
            Dictionary<string, bool> itemMaps = new Dictionary<string, bool>();  

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

                        if (!itemMaps.ContainsKey(fromItemUrl))
                        {
                            itemMaps.Add(fromItemUrl, true);
                             
                            Dictionary<string, string> moreItemRow = new Dictionary<string, string>();
                            moreItemRow.Add("detailPageUrl", fromItemUrl);
                            moreItemRow.Add("detailPageName", fromItemUrl);
                            moreItemRow.Add("itemId", fromItemId);
                            moreItemRow.Add("itemName", fromItemName);

                            moreItemEW.AddRow(moreItemRow);
                        }


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

                        this.GenerateRelatedItemFile(fromItemUrl, htmlDoc);
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

        private void GenerateRelatedItemFile(string fromItemUrl, HtmlAgilityPack.HtmlDocument htmlDoc)
        {
            HtmlNode fromTitleNode = htmlDoc.DocumentNode.SelectSingleNode("//dd[@class=\"lemmaWgt-lemmaTitle-title\"]/h1");
            string fromItemName = CommonUtil.HtmlDecode(fromTitleNode.InnerText).Trim();

            HtmlNode fromItemBaseInfoNode = htmlDoc.DocumentNode.SelectSingleNode("//div[@class=\"lemmaWgt-promotion-rightPreciseAd\"]");
            string fromItemId = fromItemBaseInfoNode.GetAttributeValue("data-lemmaid", "");
            string fromItemTitle = fromItemBaseInfoNode.GetAttributeValue("data-lemmatitle", "");


            String exportDir = this.RunPage.GetExportDir();
            string resultFilePath = Path.Combine(exportDir, "百度百科_词条关联_" + fromItemTitle + "_" + fromItemId + ".xlsx");

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

            HtmlNodeCollection aNodes = htmlDoc.DocumentNode.SelectNodes("//a");
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

        private ExcelWriter CreateItemBaseInfoWriter(int fileIndex)
        {
            String exportDir = this.RunPage.GetExportDir();
            string resultFilePath = Path.Combine(exportDir, "百度百科_朝代_基本信息_" + fileIndex.ToString() + ".xlsx");

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("url", 0);
            resultColumnDic.Add("itemId", 1);
            resultColumnDic.Add("itemName", 2);
            resultColumnDic.Add("summaryInfo", 3); 
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }         
    }
}