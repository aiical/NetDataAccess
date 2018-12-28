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

                        string beginEndYearStr = this.GetBeginEndDateTextFromSummary(summaryInfo);
                        resultRow.Add("summaryDate", beginEndYearStr);

                        string[] dateParts = this.SplitBeginEndDateStr(beginEndYearStr);
                        resultRow.Add("summaryDateBegin", dateParts == null ? "" : dateParts[0]);
                        resultRow.Add("summaryDateEnd", dateParts == null || dateParts.Length == 1 ? "" : dateParts[1]);

                        Nullable<int> summaryYearBegin = dateParts == null ? null : (Nullable<int>)this.GetYear(dateParts[0]);
                        resultRow.Add("summaryYearBegin", summaryYearBegin == null ? "" : summaryYearBegin.ToString());

                        Nullable<int> summaryYearEnd = dateParts == null || dateParts.Length == 1 ? null : (Nullable<int>)this.GetYear(dateParts[1]);
                        resultRow.Add("summaryYearEnd", summaryYearEnd == null ? "" : summaryYearEnd.ToString());

                        string birthInfo = this.GetFormattedTimeInfo(this.GetBirthTextFromSummary(summaryInfo));
                        resultRow.Add("birthInfo", birthInfo);

                        Nullable<int> birthYear = birthInfo.Length == 0 ? null : (Nullable<int>)this.GetYear(birthInfo);
                        resultRow.Add("birthYear", birthYear == null ? "" : birthYear.ToString());

                        string[] summaryShiDaiArray = this.GetShiDai(new List<string>() { sourceRow["summaryInfo"] }, shiDaiDic);
                        string summaryShiDai = summaryShiDaiArray.Length == 0 ? "" : summaryShiDaiArray[0];// CommonUtil.StringArrayToString(summaryShiDaiArray, ";");
                        resultRow.Add("summaryShiDai", summaryShiDai);

                        List<Nullable<int>> propertyYearList = new List<Nullable<int>>();

                        List<string> propertyDateBeginList = new List<string>();
                        List<string> propertyYearBeginList = new List<string>();
                        foreach (string beginProperty in this.DateBeginPropertyList)
                        {
                            string beginPropertyValue = sourceRow[beginProperty];
                            string formattedTime = this.GetFormattedTimeInfo(beginPropertyValue);
                            if (formattedTime.Length > 0)
                            {
                                propertyDateBeginList.Add(formattedTime);
                                Nullable<int> year = this.GetYear(formattedTime);
                                if (year != null)
                                {
                                    propertyYearBeginList.Add(year.ToString());
                                    propertyYearList.Add(year);
                                }
                            }
                        }
                        resultRow.Add("propertyDateBegin", CommonUtil.StringArrayToString(propertyDateBeginList.ToArray(), ";"));
                        resultRow.Add("propertyYearBegin", CommonUtil.StringArrayToString(propertyYearBeginList.ToArray(), ";"));

                        List<string> propertyDateEndList = new List<string>();
                        List<string> propertyYearEndList = new List<string>();
                        foreach (string endProperty in this.DateEndPropertyList)
                        {
                            string endPropertyValue = sourceRow[endProperty];
                            string formattedTime = this.GetFormattedTimeInfo(endPropertyValue);
                            if (formattedTime.Length > 0)
                            {
                                propertyDateEndList.Add(formattedTime);
                                Nullable<int> year = this.GetYear(formattedTime);
                                if (year != null)
                                {
                                    propertyYearEndList.Add(year.ToString());
                                    propertyYearList.Add(year);
                                }
                            }
                        }
                        resultRow.Add("propertyDateEnd", CommonUtil.StringArrayToString(propertyDateEndList.ToArray(), ";"));
                        resultRow.Add("propertyYearEnd", CommonUtil.StringArrayToString(propertyYearEndList.ToArray(), ";"));


                        List<string> propertyTexts = new List<string>();
                        foreach (string propertyName in ShiDaiPropertyList)
                        {
                            string propertyText = sourceRow[propertyName];
                            resultRow.Add(propertyName, this.ProcessDataText(propertyText));
                            propertyTexts.Add(propertyText);
                        }

                        string[] propertyShiDaiArray = this.GetShiDai(propertyTexts, shiDaiDic);
                        string propertyShiDai = propertyShiDaiArray == null || propertyShiDaiArray.Length == 0 ? "" : propertyShiDaiArray[0];
                        resultRow.Add("propertyShiDai", propertyShiDai);

                        string shiDai = propertyShiDai.Length != 0 ? propertyShiDai : summaryShiDai;
                        if (shiDai.Length > 0)
                        {
                            int[] shiDaiBeginEndYear = this.GetShiDaiBeginEndYear(shiDai, shiDaiDic);
                            Nullable<bool> summaryYearInScope = this.CheckInScope(new Nullable<int>[] { summaryYearBegin, summaryYearEnd, birthYear }, shiDaiBeginEndYear);

                            resultRow.Add("summaryYearInScope", summaryYearInScope == null ? "" : ((bool)summaryYearInScope ? "1" : "0"));
                        }

                        string propertyYearMatchedShiDai = this.MatchShiDai(propertyYearList, shiDaiDic);
                        resultRow.Add("propertyYearMatchedShiDai", propertyYearMatchedShiDai);

                        Nullable<int> beginYear = null;
                        Nullable<int> endYear = null;
                        if ((propertyYearBeginList != null && propertyYearBeginList.Count > 0) || summaryYearBegin != null || birthYear != null)
                        {
                            if (propertyYearEndList != null && propertyYearBeginList.Count > 0)
                            {
                                beginYear = int.Parse(propertyYearBeginList[0]);
                            }
                            else if (birthYear != null)
                            {
                                beginYear = birthYear;
                            }
                            else
                            {
                                beginYear = summaryYearBegin;
                            }
                        }
                        if ((propertyYearEndList != null && propertyYearEndList.Count > 0) || summaryYearEnd != null)
                        {
                            if (propertyYearEndList != null && propertyYearEndList.Count > 0)
                            {
                                endYear = int.Parse(propertyYearEndList[0]);
                            }
                            else
                            {
                                endYear = summaryYearEnd;
                            }
                        }
                        if (beginYear == null && endYear == null&& shiDai.Length>0)
                        {
                            int[] shiDaiYears = shiDaiDic[shiDai];
                            beginYear = shiDaiYears[0];
                            endYear = shiDaiYears[1];
                        }

                        resultRow.Add("beginYear", beginYear.ToString());
                        resultRow.Add("endYear", endYear.ToString());


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

        private Nullable<bool> CheckInScope(Nullable<int>[] years, int[] scope)
        {
            Nullable<bool> isInScope = null;
            if (years != null && scope != null && scope.Length != 0)
            {
                foreach (Nullable<int> year in years)
                {
                    if (year == null)
                    {
                        //不做处理
                    }
                    else
                    {
                        //允许100年的误差
                        if (scope[0] < year + 100 && scope[1] > year - 100)
                        {
                            isInScope = true;
                        }
                        else
                        {
                            return false;
                        }
                    }
                }
            }
            return isInScope;
        }

        private Nullable<int> GetYear(string formatedDateDescription)
        {
            if (formatedDateDescription.Length > 0)
            {
                string[] partDateStrs = formatedDateDescription.Split(new string[] { "||" }, StringSplitOptions.RemoveEmptyEntries);
                int year = 0;
                bool isBC = false;
                foreach (string partDateStr in partDateStrs)
                {
                    string p = partDateStr.Replace("|", "");
                    if (p == "前")
                    {
                        isBC = true;
                    }
                    else if (p.Contains("年"))
                    {
                        string yearStr = p.Replace("年", "").Trim();
                        int.TryParse(yearStr, out year);
                    }
                }
                return year == 0 ? null : (Nullable<int>)(isBC ? (0 - year) : year);
            }
            return null;
        }

        private int[] GetShiDaiBeginEndYear(string shiDai,Dictionary<string,int[]> shiDaiDic)
        {
            if (shiDai.Length > 0)
            {
                if (shiDaiDic.ContainsKey(shiDai))
                {
                    return shiDaiDic[shiDai];
                }
            }
            return null;
        }

        private string[] _DateBeginEndSpliter = null;
        private string[] DateBeginEndSpliter
        {
            get
            {
                if (this._DateBeginEndSpliter == null)
                {
                    this._DateBeginEndSpliter = new string[] { "----", "--", "——", "~", "一", "～", "-", "—", "－", "─", "―", "至" };
                }
                return this._DateBeginEndSpliter;
            }
        }
        private string[] SplitBeginEndDateStr(string yearInfo)
        {
            if (yearInfo.Length > 0)
            {
                string[] dateParts = yearInfo.Split(this.DateBeginEndSpliter, StringSplitOptions.None);
                if (dateParts.Length == 2)
                {
                    string[] formattedYearParts = new string[2];
                    for (int i = 0; i < dateParts.Length; i++)
                    {
                        formattedYearParts[i] = this.GetFormattedTimeInfo(dateParts[i]);
                    }

                    return formattedYearParts;
                }
                else
                {
                    string[] formattedDateParts = new string[1] { this.GetFormattedTimeInfo(dateParts[0]) };
                    return formattedDateParts;
                }
            }
            return null;
        }

        private string GetFormattedTimeInfo(string timeStr)
        {
            if (timeStr.Length > 0)
            {
                StringBuilder resultStr = new StringBuilder();

                Dictionary<string,string> checkStrDic = new Dictionary<string,string>();
                checkStrDic.Add("约","约");
                checkStrDic.Add("公元前", "前");
                checkStrDic.Add("初", "");
                checkStrDic.Add("公元", "");
                checkStrDic.Add("前", "前");
                checkStrDic.Add("农历" ,"农历");
                foreach (string checkStr in checkStrDic.Keys)
                {
                    if (timeStr.Contains(checkStr))
                    {
                        timeStr = timeStr.Replace(checkStr, "");
                        string subStr = checkStrDic[checkStr];
                        if (subStr.Length > 0)
                        {
                            resultStr.Append("|" + subStr + "|");
                        }
                    }
                }

                string[] removeStrs = new string[] { "?", "？" };
                foreach (string removeStr in removeStrs)
                {
                    if (timeStr.Contains(removeStr))
                    {
                        timeStr = timeStr.Replace(removeStr, "");
                    }
                }

                Dictionary<string, string> chineseToNumDic = new Dictionary<string, string>();
                chineseToNumDic.Add("１", "1");
                chineseToNumDic.Add("２", "2");
                chineseToNumDic.Add("３", "3");
                chineseToNumDic.Add("４", "4");
                chineseToNumDic.Add("５", "5");
                chineseToNumDic.Add("６", "6");
                chineseToNumDic.Add("７", "7");
                chineseToNumDic.Add("８", "8");
                chineseToNumDic.Add("９", "9");
                chineseToNumDic.Add("０", "0");
                chineseToNumDic.Add("一", "1");
                chineseToNumDic.Add("二", "2");
                chineseToNumDic.Add("三", "3");
                chineseToNumDic.Add("四", "4");
                chineseToNumDic.Add("五", "5");
                chineseToNumDic.Add("六", "6");
                chineseToNumDic.Add("七", "7");
                chineseToNumDic.Add("八", "8");
                chineseToNumDic.Add("九", "9");
                chineseToNumDic.Add("零", "0");
                chineseToNumDic.Add("〇", "0");
                chineseToNumDic.Add("o", "0");
                chineseToNumDic.Add("O", "0");
                chineseToNumDic.Add("寒月", "10月");
                chineseToNumDic.Add("冬月", "11月");
                chineseToNumDic.Add("腊月", "12月");
                chineseToNumDic.Add("正月", "1月");
                foreach (string chinese in chineseToNumDic.Keys)
                {
                    timeStr = timeStr.Replace(chinese, chineseToNumDic[chinese]);
                }
                
                //处理“十”
                int tenStrIndex = timeStr.IndexOf("十");
                while (tenStrIndex > -1)
                {
                    if (tenStrIndex == 0)
                    {
                        timeStr = "1" + timeStr.Substring(tenStrIndex + 1);
                    }
                    else 
                    {
                        char preChar = timeStr[tenStrIndex-1];
                        if (preChar >= '0' && preChar <= '9')
                        {
                            timeStr = timeStr.Substring(0, tenStrIndex) + timeStr.Substring(tenStrIndex + 1);
                        }
                        else
                        {
                            timeStr = timeStr.Substring(0, tenStrIndex) + "1" + timeStr.Substring(tenStrIndex + 1);
                        }
                    }
                    tenStrIndex = timeStr.IndexOf("十");
                }

                int yearCharIndex = timeStr.IndexOf("年");
                if (yearCharIndex > -1)
                {
                    int tempIndex = yearCharIndex-1;
                    while (tempIndex > -1)
                    {
                        char tempChar = timeStr[tempIndex];
                        if (tempChar >= '0' && tempChar <= '9')
                        {
                            tempIndex = tempIndex - 1;
                        }
                        else
                        {
                            break;
                        }
                    }
                    timeStr = timeStr.Substring(tempIndex + 1);
                }

                if (timeStr.Contains(".") || timeStr.Contains("．") || timeStr.Contains("-"))
                {
                    string[] timePartStrs = timeStr.Split(new string[] { ".", "．", "-" }, StringSplitOptions.RemoveEmptyEntries);
                    if (timePartStrs.Length > 0)
                    {
                        resultStr.Append("|" + timePartStrs[0] + "年|");
                    }
                    if (timePartStrs.Length > 1)
                    {
                        resultStr.Append("|" + timePartStrs[1] + "月|");
                    }
                    if (timePartStrs.Length > 2)
                    {
                        resultStr.Append("|" + timePartStrs[2] + "日|");
                    }
                }
                else
                {
                    List<string> spliterStrs = new List<string>() { "年", "月", "日" };
                    foreach (string spliterStr in spliterStrs)
                    {
                        int num = 0;
                        if (timeStr.Contains(spliterStr))
                        {
                            string[] yearSplitParts = timeStr.Split(new string[] { spliterStr }, StringSplitOptions.None);
                            if (int.TryParse(yearSplitParts[0], out num))
                            {
                                resultStr.Append("|" + num.ToString() + spliterStr + "|");
                            }
                            timeStr = yearSplitParts.Length > 1 ? yearSplitParts[1] : "";
                        }
                        else
                        {
                            if (int.TryParse(timeStr, out num))
                            {
                                resultStr.Append("|" + num.ToString() + spliterStr + "|");
                                break;
                            }
                        }
                    }
                }
                return resultStr.ToString();
            }
            else
            {
                return "";
            }
        }

        private string[] GetShiDai(List<string> sourceTextList, Dictionary<string, int[]> shiDaiDic)
        {
            Dictionary<string, bool> matchedShiDaiDic = new Dictionary<string, bool>();
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
            List<string> matchedShiDaiList = new List<string>();
            foreach (string shiDai in matchedShiDaiDic.Keys)
            {
                matchedShiDaiList.Add(shiDai);
            }
            return matchedShiDaiList.ToArray();
        }
        private string MatchShiDai(List<Nullable<int>> yearList, Dictionary<string, int[]> shiDaiDic)
        {
            foreach (Nullable<int> year in yearList)
            {
                if (year != null)
                {
                    int yearInt = (int)year;
                    foreach (string shiDai in shiDaiDic.Keys)
                    {
                        int[] scope = shiDaiDic[shiDai];
                        if (yearInt >= scope[0] && yearInt <= scope[1])
                        {
                            return shiDai;
                        }
                    }
                }
            }
            return "";
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

        private string GetBeginEndDateTextFromSummary(string summaryText)
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
                        string dateText = summaryText.Substring(matchedBeginIndex + 1, matchedEndIndex - matchedBeginIndex - 1);
                        bool got = false;
                        string[] partDateTexts = dateText.Split(splitSymbolArray, StringSplitOptions.RemoveEmptyEntries);
                        for (int i = 0; i < partDateTexts.Length; i++)
                        {
                            string partDateText = partDateTexts[i];
                            if (partDateText.Length > 2 && Regex.IsMatch(partDateText, ".*[0-9]|[0-9].*"))
                            {
                                return partDateText;
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
            List<string> removeChars = new List<string>() { " ", "　"};
            foreach (string removeChar in removeChars)
            {
                summaryText = summaryText.Replace(removeChar, "");
            }
            
            List<string> birthTextList = new List<string>() { };
            birthTextList.Add("年出生");
            birthTextList.Add("月出生");
            birthTextList.Add("日出生");
            birthTextList.Add("出生");
            birthTextList.Add("年生");
            birthTextList.Add("月生");
            birthTextList.Add("日生");
            Dictionary<char, bool> textSplitSymbolDic = new Dictionary<char, bool>();
            textSplitSymbolDic.Add(',', true);
            textSplitSymbolDic.Add('。', true);
            textSplitSymbolDic.Add('，', true);
            textSplitSymbolDic.Add('、', true);
            textSplitSymbolDic.Add('\n', true);


            foreach (string birthText in birthTextList)
            {
                int partBeginIndex = -1;
                int dateTextBeginIndex = summaryText.IndexOf(birthText);
                if (dateTextBeginIndex > -1)
                {
                    for (int i = dateTextBeginIndex - 1; i >= 0; i--)
                    {
                        char checkChar = summaryText[i];
                        if (textSplitSymbolDic.ContainsKey(checkChar))
                        {
                            partBeginIndex = i;
                            break;
                        }
                    }
                    return summaryText.Substring(partBeginIndex + 1, dateTextBeginIndex - partBeginIndex - 1).Trim()+ birthText;
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
        private List<string> _DateBeginPropertyList = null;
        private List<string> DateBeginPropertyList
        {
            get
            {
                if (this._DateBeginPropertyList == null)
                {
                    this._DateBeginPropertyList = new List<string>() { 
                        "出生日期",
                        "出生时间"};
                }
                return this._DateBeginPropertyList;
            }
        }
        private List<string> _DateEndPropertyList = null;
        private List<string> DateEndPropertyList
        {
            get
            {
                if (this._DateEndPropertyList == null)
                {
                    this._DateEndPropertyList = new List<string>() { 
                        "去世日期",
                        "去世时间", 
                        "逝世日期", 
                        "逝世时间" };
                }
                return this._DateEndPropertyList;
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
            resultColumnDic.Add("summaryDate", 3);
            resultColumnDic.Add("summaryDateBegin", 4);
            resultColumnDic.Add("summaryDateEnd", 5);
            resultColumnDic.Add("summaryYearBegin", 6);
            resultColumnDic.Add("summaryYearEnd", 7);
            resultColumnDic.Add("birthInfo", 8);
            resultColumnDic.Add("birthYear", 9);
            resultColumnDic.Add("summaryShiDai", 10);
            resultColumnDic.Add("propertyDateBegin", 11);
            resultColumnDic.Add("propertyDateEnd", 12);
            resultColumnDic.Add("propertyYearBegin", 13);
            resultColumnDic.Add("propertyYearEnd", 14);
            resultColumnDic.Add("propertyShiDai", 15);
            resultColumnDic.Add("summaryInfo", 16);
            resultColumnDic.Add("summaryYearInScope", 17);
            resultColumnDic.Add("propertyYearMatchedShiDai", 18);
            resultColumnDic.Add("beginYear", 19);
            resultColumnDic.Add("endYear", 20);
            for (int i = 0; i < this.ShiDaiPropertyList.Count; i++)
            {
                resultColumnDic.Add(this.ShiDaiPropertyList[i], 21 + i);
            }
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }
    }
}
