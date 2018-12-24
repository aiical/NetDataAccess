using HtmlAgilityPack;
using NetDataAccess.Base.Common;
using NetDataAccess.Base.Config;
using NetDataAccess.Base.DB;
using NetDataAccess.Base.Reader;
using NetDataAccess.Base.UI;
using NetDataAccess.Base.Writer;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace NetDataAccess.Extended.LiShi.BaiDuBaiKeRenWu
{
    public class ProcessLiShiRenWuShiDaiInfo : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            this.GetRenWuShiDaiDataInfos(listSheet);
            return true;
        }

        private void GetRenWuShiDaiDataInfos(IListSheet listSheet)
        {
            string[] parameters = this.Parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            string renWuShiDaiInfoFilePath = parameters[0];
            string shiDaiInfoFilePath = parameters[1];
            string exportFilePath = parameters[2];


            try
            {
                ExcelReader shiDaiInfoReader = new ExcelReader(shiDaiInfoFilePath, "时代");
                Dictionary<string, int[]> shiDaiDic = new Dictionary<string, int[]>();
                int shiDaiRowCount = shiDaiInfoReader.GetRowCount();
                for (int i = 0; i < shiDaiRowCount; i++)
                {
                    Dictionary<string, string> shiDaiRow = shiDaiInfoReader.GetFieldValues(i);
                    string shiDaiNameStr = shiDaiRow["时期"];
                    int beginYear = int.Parse(shiDaiRow["起始年份"]);
                    int endYear = int.Parse(shiDaiRow["终止年份"]);
                    string[] shiDaiNames = shiDaiNameStr.Split(new string[] { "、"}, StringSplitOptions.RemoveEmptyEntries);
                    for (int j = 0; j < shiDaiNames.Length; j++)
                    {
                        string shiDaiName = shiDaiNames[j].Trim();
                        if (!shiDaiDic.ContainsKey(shiDaiName))
                        {
                            shiDaiDic.Add(shiDaiName, new int[] { beginYear, endYear });
                        }
                        else
                        {
                            throw new Exception("已经存在了时期, name = " + shiDaiName);
                        }
                    }
                }

                ExcelReader renWuInfoReader = new ExcelReader(renWuShiDaiInfoFilePath);
                int sourceRowCount = renWuInfoReader.GetRowCount();

                ExcelWriter resultWriter = this.CreatePropertyVaueWriter(exportFilePath);
                for (int i = 0; i < sourceRowCount; i++)
                {
                    Dictionary<string, string> sourceRow = renWuInfoReader.GetFieldValues(i);
                    try
                    {
                        Dictionary<string, string> resultRow = new Dictionary<string, string>();
                        resultRow.Add("url", sourceRow["url"]);
                        resultRow.Add("itemId", sourceRow["itemId"]);
                        resultRow.Add("itemName", sourceRow["itemName"]);

                        
                        string summaryInfo = sourceRow["summaryInfo"];
                        resultRow.Add("summaryInfo", summaryInfo);

                        string beginEndYearStr = this.GetBeginEndYearTextFromSummary(summaryInfo);
                        resultRow.Add("summaryYear", beginEndYearStr);

                        string[] yearParts = this.SplitBeginEndYearStr(beginEndYearStr);
                        resultRow.Add("summaryYearBegin", yearParts == null ? "" : yearParts[0]);
                        resultRow.Add("summaryYearEnd", yearParts == null ? "" : yearParts[1]);

                        string birthInfo = this.GetBirthTextFromSummary(summaryInfo);
                        resultRow.Add("birthInfo", birthInfo);

                        string summaryShiDai= CommonUtil.StringArrayToString(this.GetShiDai(new List<string>() { sourceRow["summaryInfo"] }, shiDaiDic), ";");
                        resultRow.Add("summaryShiDai", summaryShiDai);

                        string propertyYearBegin = "";
                        foreach (string beginProperty in this.YearBeginPropertyList)
                        {

                        }
                        resultRow.Add("propertyYearBegin", propertyYearBegin);

                        string propertyYearEnd = "";
                        foreach (string endProperty in this.YearEndPropertyList)
                        {

                        }
                        resultRow.Add("propertyYearEnd", propertyYearEnd);


                        List<string> propertyTexts = new List<string>();
                        foreach (string propertyName in ShiDaiPropertyList)
                        {
                            string propertyText = sourceRow[propertyName];
                            resultRow.Add(propertyName, this.ProcessDataText(propertyText));
                            propertyTexts.Add(propertyText);
                        }

                        resultRow.Add("propertyShiDai", CommonUtil.StringArrayToString(this.GetShiDai(propertyTexts, shiDaiDic), ";")); 
                        resultWriter.AddRow(resultRow);
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
                resultWriter.SaveToDisk(); 

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private string[] _YearBeginEndSpliter = null;
        private string[] YearBeginEndSpliter
        {
            get
            {
                if (this._YearBeginEndSpliter == null)
                {
                    this._YearBeginEndSpliter = new string[] { "----", "--", "——", "~", "一", "～", "-", "—" };
                }
                return this._YearBeginEndSpliter;
            }
        }
        private string[] SplitBeginEndYearStr(string yearInfo)
        {
            if (yearInfo.Length > 0)
            {
                string[] yearParts = yearInfo.Split(this.YearBeginEndSpliter, StringSplitOptions.RemoveEmptyEntries);
                if (yearParts.Length == 2)
                {
                    string[] formattedYearParts = new string[2];
                    for (int i = 0; i < yearParts.Length; i++)
                    {
                        formattedYearParts[i] = yearParts[i].Trim();
                    }

                    return formattedYearParts;
                }                
            }
            return null;
        }

        private string[] GetShiDai(List<string> sourceTextList, Dictionary<string, int[]> shiDaiDic)
        {
            Dictionary<string, bool> matchedShiDaiDic = new Dictionary<string,bool>();
            foreach (string sourceText in sourceTextList)
            {
                if (sourceText.Length > 0)
                {
                    foreach (string shiDai in shiDaiDic.Keys)
                    {
                        if (!matchedShiDaiDic.ContainsKey(shiDai) && sourceText.Contains(shiDai))
                        {
                            matchedShiDaiDic.Add(shiDai, true);
                        }
                    }
                }
            }
            List<string> matchedShiDaiList= new List<string>();
            foreach (string shiDai in matchedShiDaiDic.Keys)
            {
                matchedShiDaiList.Add(shiDai);
            }
            return matchedShiDaiList.ToArray();
        }

        private string ProcessDataText(string sourceText)
        {
            sourceText = sourceText.Trim();
            string[] removeChars = new string[] { "（", "）", "(", ")", "/", "[", "]", "【", "】", "、", "，", ":", "：", "“", "”", "。", " ", " " };

            foreach (string removeChar in removeChars)
            {
              sourceText =   sourceText.Replace(removeChar, "");
            }

            string[] deleteOnlyStrings = new string[] { "?", "??", "?年", "？", "？年", "-", "--", "---", "----", "—", "——" };
            foreach (string deleteOnlyString in deleteOnlyStrings)
            {
                if (deleteOnlyString == sourceText)
                {
                    return "";
                }
            }
            return sourceText;
        }

        private string GetBeginEndYearTextFromSummary(string summaryText)
        {
            int summaryLength = summaryText.Length;
            int checkIndex = 0;
            int matchedBeginIndex = -1;
            int matchedEndIndex = -1;

            string[] splitSymbolArray = new string[] { ";", ",", "；", "，", "[", "]" };

            Dictionary<char, bool> matchBeginCharDic = new Dictionary<char, bool>();
            matchBeginCharDic.Add('(', true);
            matchBeginCharDic.Add('（', true);

            Dictionary<char, bool> matchEndCharDic = new Dictionary<char, bool>();
            matchEndCharDic.Add('）', true);
            matchEndCharDic.Add(')', true);

            bool inMatched = false;
            while (checkIndex < summaryLength)
            {
                char checkChar = summaryText[checkIndex];
                if (inMatched)
                {
                    if (matchEndCharDic.ContainsKey(checkChar))
                    {
                        matchedEndIndex = checkIndex;
                        inMatched = false;
                        string yearText = summaryText.Substring(matchedBeginIndex + 1, matchedEndIndex - matchedBeginIndex - 1);
                        bool got = false;
                        string[] partYearTexts = yearText.Split(splitSymbolArray, StringSplitOptions.RemoveEmptyEntries);
                        for (int i = 0; i < partYearTexts.Length; i++)
                        {
                            string partYearText = partYearTexts[i];
                            if (partYearText.Length > 2 && Regex.IsMatch(partYearText, ".*[0-9]|[0-9].*"))
                            {
                                return partYearText;
                            }
                        }
                        if (!got)
                        {
                            matchedBeginIndex = -1;
                            matchedEndIndex = -1;
                        }
                    }
                }
                else
                {
                    if (matchBeginCharDic.ContainsKey(checkChar))
                    {
                        matchedBeginIndex = checkIndex;
                        inMatched = true;
                    }
                }

                checkIndex++;
            } 
            return "";
        }
        private string GetBirthTextFromSummary(string summaryText)
        { 
            List<string> birthTextList = new List<string>() { };
            birthTextList.Add("年出生");
            birthTextList.Add("月出生");
            birthTextList.Add("日出生");
            birthTextList.Add("年生");
            birthTextList.Add("月生");
            birthTextList.Add("日生");
            Dictionary<char, bool> textSplitSymbolDic = new Dictionary<char, bool>();
            textSplitSymbolDic.Add(',', true);
            textSplitSymbolDic.Add('。', true);
            textSplitSymbolDic.Add('，', true);
            textSplitSymbolDic.Add('\n', true);


            foreach (string birthText in birthTextList)
            {
                int partBeginIndex = -1;
                int yearTextBeginIndex = summaryText.IndexOf(birthText);
                if (yearTextBeginIndex > -1)
                {
                    for (int i = yearTextBeginIndex - 1; i >= 0; i--)
                    {
                        char checkChar = summaryText[i];
                        if (textSplitSymbolDic.ContainsKey(checkChar))
                        {
                            partBeginIndex = i;
                            break;
                        }
                    }
                    return summaryText.Substring(partBeginIndex + 1, yearTextBeginIndex - partBeginIndex - 1).Trim()+ birthText;
                }
            }
            return "";
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
        private List<string> _ShiDaiPropertyList = null;
        private List<string> ShiDaiPropertyList
        {
            get
            {
                if (this._ShiDaiPropertyList == null)
                {
                    this._ShiDaiPropertyList = new List<string>() {
                        "所处时代", 
                        "日期", 
                        "时间", 
                        "时期", 
                        "时代", 
                        "年代", 
                        "国家", 
                        "国籍", 
                        "朝代", 
                        "出生日期",
                        "出生时间", 
                        "去世日期",
                        "去世时间", 
                        "逝世日期", 
                        "逝世时间" };
                }
                return this._ShiDaiPropertyList;
            }
        }
        private List<string> _YearBeginPropertyList = null;
        private List<string> YearBeginPropertyList
        {
            get
            {
                if (this._YearBeginPropertyList == null)
                {
                    this._YearBeginPropertyList = new List<string>() { 
                        "出生日期",
                        "出生时间"};
                }
                return this._YearBeginPropertyList;
            }
        }
        private List<string> _YearEndPropertyList = null;
        private List<string> YearEndPropertyList
        {
            get
            {
                if (this._YearEndPropertyList == null)
                {
                    this._YearEndPropertyList = new List<string>() { 
                        "去世日期",
                        "去世时间", 
                        "逝世日期", 
                        "逝世时间" };
                }
                return this._YearEndPropertyList;
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

        private ExcelWriter CreatePropertyVaueWriter(string resultFilePath)
        { 
            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("url", 0);
            resultColumnDic.Add("itemId", 1);
            resultColumnDic.Add("itemName", 2);
            resultColumnDic.Add("summaryYear", 3);
            resultColumnDic.Add("summaryYearBegin", 4);
            resultColumnDic.Add("summaryYearEnd", 5);
            resultColumnDic.Add("birthInfo", 6);
            resultColumnDic.Add("summaryShiDai", 7);
            resultColumnDic.Add("propertyYearBegin", 8);
            resultColumnDic.Add("propertyYearEnd", 9);
            resultColumnDic.Add("propertyShiDai", 10);
            resultColumnDic.Add("summaryInfo", 11);
            for (int i = 0; i < this.ShiDaiPropertyList.Count; i++)
            {
                resultColumnDic.Add(this.ShiDaiPropertyList[i], 11 + i);
            }
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }
    }
}
