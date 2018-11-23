using HtmlAgilityPack;
using NetDataAccess.Base.Browser;
using NetDataAccess.Base.Common;
using NetDataAccess.Base.Config;
using NetDataAccess.Base.DB;
using NetDataAccess.Base.EnumTypes;
using NetDataAccess.Base.Reader;
using NetDataAccess.Base.UI;
using NetDataAccess.Base.Web;
using NetDataAccess.Base.Writer;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace NetDataAccess.Extended.LunWen.All
{ 
    public class ProcessAllKeywords : ExternalRunWebPage
    { 
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            this.GetYearWordsMatrixCount(listSheet);
            //this.GetYearWordsCount(listSheet);
            return true;
        }
        private void GetYearWordsMatrixCount(IListSheet listSheet)
        {
            try
            {
                string[] parameters = this.Parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                string allKeywordsFilePath = parameters[0];
                string exportDirPath = parameters[2];


                ExcelReader er = new ExcelReader(allKeywordsFilePath);
                int inputRowCount = er.GetRowCount();


                List<string> keywordList = new List<string>();
                Dictionary<int, Dictionary<string, List<string>>> yearSourceWordList = new Dictionary<int, Dictionary<string, List<string>>>();
                for (int i = 0; i < inputRowCount; i++)
                {
                    Dictionary<string, string> row = er.GetFieldValues(i);
                    string source = row["source"];
                    int year = int.Parse( row["year"]);
                    string keyword = row["keyword"];

                    if (!keywordList.Contains(keyword))
                    {
                        keywordList.Add(keyword);
                    }

                    if (!yearSourceWordList.ContainsKey(year))
                    { 
                        yearSourceWordList.Add(year, new Dictionary<string, List<string>>());
                    }
                    Dictionary<string, List<string>> sourceWordList = yearSourceWordList[year];
                    if (!sourceWordList.ContainsKey(source))
                    {
                        sourceWordList.Add(source, new List<string>());
                    }
                    List<string> wordList = sourceWordList[source];
                    if (!wordList.Contains(keyword))
                    {
                        wordList.Add(keyword);
                    }
                }
                 
                Dictionary<string, Dictionary<string, int>> totalYearMartixDataDic = new Dictionary<string, Dictionary<string, int>>(); 

                foreach (int year in yearSourceWordList.Keys)
                {
                    Dictionary<string, Dictionary<string, int>> yearMatrixDataDic = new Dictionary<string, Dictionary<string, int>>(); 
                    Dictionary<string, List<string>> sourceWordList = yearSourceWordList[year];
                    foreach (string source in sourceWordList.Keys)
                    {
                        List<string> kwList = sourceWordList[source];
                        for (int i = 0; i < kwList.Count; i++)
                        {
                            string kw_i = kwList[i]; 
                            if (!yearMatrixDataDic.ContainsKey(kw_i))
                            {
                                yearMatrixDataDic.Add(kw_i, new Dictionary<string, int>());
                            }
                            Dictionary<string, int> iDic = yearMatrixDataDic[kw_i];

                            if (!iDic.ContainsKey(kw_i))
                            {
                                iDic.Add(kw_i, 1);
                            }
                            else
                            {
                                iDic[kw_i] = iDic[kw_i] + 1;
                            }

                            /*
                            if (!totalYearMartixDataDic.ContainsKey(kw_i))
                            {
                                totalYearMartixDataDic.Add(kw_i, new Dictionary<string, int>());
                            }
                            Dictionary<string, int> iTotalDic = totalYearMartixDataDic[kw_i];
                            if (!iTotalDic.ContainsKey(kw_i))
                            {
                                iTotalDic.Add(kw_i, 1);
                            }
                            else
                            {
                                iTotalDic[kw_i] = iTotalDic[kw_i] + 1;
                            }
                             * */

                            for (int j = 0; j < kwList.Count; j++)
                            {
                                string kw_j = kwList[j];
                                if (kw_i != kw_j)
                                {
                                    if (!iDic.ContainsKey(kw_j))
                                    {
                                        iDic.Add(kw_j, 1);
                                    }
                                    else
                                    {
                                        iDic[kw_j] = iDic[kw_j] + 1;
                                    }
                                    
                                    /*
                                    if (!iTotalDic.ContainsKey(kw_j))
                                    {
                                        iTotalDic.Add(kw_j, 1);
                                    }
                                    else
                                    {
                                        iTotalDic[kw_i] = iTotalDic[kw_j] + 1;
                                    }
                                     */
                                }
                            }
                        }
                    }
                     
                    CsvWriter resultWriter = this.GetMatrixCsvWriter(exportDirPath, year, keywordList); 

                    for (int i = 0; i < keywordList.Count; i++)
                    {
                        Dictionary<string, string> matrixRow = new Dictionary<string, string>();
                        string kw_i = keywordList[i];
                        matrixRow["keywordMatrix"] = kw_i;
                        Dictionary<string, int> iMatrixDataDic = yearMatrixDataDic.ContainsKey(kw_i) ? yearMatrixDataDic[kw_i] : null;
                        for (int j = 0; j < keywordList.Count; j++)
                        {
                            string kw_j = keywordList[j];
                            if (iMatrixDataDic == null)
                            {
                                matrixRow.Add(kw_j, "0");
                            }
                            else
                            {
                                matrixRow.Add(kw_j, iMatrixDataDic.ContainsKey(kw_j) ? iMatrixDataDic[kw_j].ToString() : "0");
                            }
                        }
                        resultWriter.AddRow(matrixRow);
                    }
                    resultWriter.SaveToDisk(); 

                    foreach (string kw_i in yearMatrixDataDic.Keys)
                    {
                        if (!totalYearMartixDataDic.ContainsKey(kw_i))
                        {
                            totalYearMartixDataDic.Add(kw_i, new Dictionary<string, int>());
                        }
                        Dictionary<string, int> iTotalDataDic = totalYearMartixDataDic[kw_i];
                        Dictionary<string, int> iDataDic = yearMatrixDataDic[kw_i];
                        foreach (string kw_j in iDataDic.Keys)
                        {
                            if (!iTotalDataDic.ContainsKey(kw_j))
                            {
                                iTotalDataDic.Add(kw_j, iDataDic[kw_j]);
                            }
                            else
                            {
                                iTotalDataDic[kw_j] = iTotalDataDic[kw_j] + iDataDic[kw_j];
                            }
                        }
                    }
                }

                CsvWriter totalRresultWriter = this.GetMatrixCsvWriter(exportDirPath, 0, keywordList);

                for (int i = 0; i < keywordList.Count; i++)
                {
                    Dictionary<string, string> matrixRow = new Dictionary<string, string>();
                    string kw_i = keywordList[i];
                    matrixRow["keywordMatrix"] = kw_i; 
                    Dictionary<string, int> iMatrixDataDic = totalYearMartixDataDic.ContainsKey(kw_i) ? totalYearMartixDataDic[kw_i] : null;
                    for (int j = 0; j < keywordList.Count; j++)
                    {
                        string kw_j = keywordList[j];
                        if (iMatrixDataDic == null)
                        {
                            matrixRow.Add(kw_j, "0");
                        }
                        else
                        {
                            matrixRow.Add(kw_j, iMatrixDataDic.ContainsKey(kw_j) ? iMatrixDataDic[kw_j].ToString() : "0");
                        }
                    }
                    totalRresultWriter.AddRow(matrixRow);
                }
                totalRresultWriter.SaveToDisk();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void GetYearWordsCount(IListSheet listSheet)
        {
            try
            {
                string[] parameters = this.Parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                string allKeywordsFilePath = parameters[0];
                string exportFilePath = parameters[1];


                ExcelReader er = new ExcelReader(allKeywordsFilePath);
                int inputRowCount = er.GetRowCount();

                Dictionary<int, Dictionary<string, int>> allYearDic = new Dictionary<int, Dictionary<string, int>>();
                Dictionary<string, int> totalYearDic = new Dictionary<string, int>();

                allYearDic.Add(0, totalYearDic);
                for (int i = 0; i < inputRowCount; i++)
                {
                    Dictionary<string, string> row = er.GetFieldValues(i);

                    string keyword = row["keyword"];
                    if (totalYearDic.ContainsKey(keyword))
                    {
                        totalYearDic[keyword] = totalYearDic[keyword] + 1;
                    }
                    else
                    {
                        totalYearDic.Add(keyword, 1);
                    }

                    int year = int.Parse(row["year"]);

                    if (!allYearDic.ContainsKey(year))
                    {
                        allYearDic.Add(year, new Dictionary<string, int>());
                    }
                    Dictionary<string, int> thisYearDic = allYearDic[year];

                    if (thisYearDic.ContainsKey(keyword))
                    {
                        thisYearDic[keyword] = thisYearDic[keyword] + 1;
                    }
                    else
                    {
                        thisYearDic.Add(keyword, 1);
                    }
                }

                DataTable dt = new DataTable();
                dt.Columns.Add("kw", typeof(string));
                dt.Columns.Add("cn", typeof(int));
                foreach (string keyword in totalYearDic.Keys)
                {
                    int totalCount = totalYearDic[keyword];
                    DataRow row = dt.NewRow();
                    row["kw"] = keyword;
                    row["cn"] = totalCount;
                    dt.Rows.Add(row);
                }
                DataRow[] sortedTotalRows = dt.Select("", "cn desc");

                List<int> years = new List<int>();
                for (int i = 2020; i > 1900; i--)
                {
                    if (allYearDic.ContainsKey(i))
                    {
                        years.Add(i);
                    }
                }

                ExcelWriter resultWriter = this.GetExcelWriter(exportFilePath, years);

                for (int i = 0; i < sortedTotalRows.Length; i++)
                {
                    DataRow totalRow = sortedTotalRows[i];
                    string keyword = totalRow["kw"].ToString();
                    int totalCount = (int)totalRow["cn"];


                    Dictionary<string, object> resultRow = new Dictionary<string, object>();
                    resultRow.Add("keyword", keyword);
                    resultRow.Add("total", totalCount);
                    foreach (int year in years)
                    {
                        Dictionary<string, int> tempYearDic = allYearDic[year];
                        if (tempYearDic.ContainsKey(keyword))
                        {
                            resultRow.Add(year.ToString(), tempYearDic[keyword]);
                        }
                        else
                        {
                            resultRow.Add(year.ToString(), 0);
                        }
                    }
                    resultWriter.AddRow(resultRow);
                }
                resultWriter.SaveToDisk();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private CsvWriter GetMatrixCsvWriter(string exportDirPath, int year, List<string> keywordList)
        {
            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            Dictionary<string, string> columnFormats = new Dictionary<string, string>();
            resultColumnDic.Add("keywordMatrix", 0);
            for (int i = 0; i < keywordList.Count; i++)
            {
                resultColumnDic.Add(keywordList[i].ToString(), i + 1);
                columnFormats.Add(keywordList[i].ToString(), "#0");
            }
            string filePath = Path.Combine(exportDirPath, year + ".csv");
            CsvWriter resultEW = new CsvWriter(filePath, resultColumnDic);
            return resultEW;
        }

        private  ExcelWriter GetMatrixExcelWriter(string exportDirPath, int year, List<string> keywordList)
        {
            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            Dictionary<string, string> columnFormats = new Dictionary<string, string>();
            resultColumnDic.Add("keywordMatrix", 0);
            for (int i = 0; i < keywordList.Count; i++)
            {
                resultColumnDic.Add(keywordList[i].ToString(), i + 1);
                columnFormats.Add(keywordList[i].ToString(), "#0");
            }
            string filePath = Path.Combine(exportDirPath, year + ".xlsx");
            ExcelWriter resultEW = new ExcelWriter(filePath, "List", resultColumnDic, columnFormats);
            return resultEW;
        }

        private ExcelWriter GetExcelWriter(string filePath, List<int> years)
        {
            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            Dictionary<string, string> columnFormats = new Dictionary<string, string>();
            resultColumnDic.Add("keyword", 0);
            resultColumnDic.Add("total", 1);
            columnFormats.Add("total", "#0");
            for (int i = 0; i < years.Count; i++)
            {
                resultColumnDic.Add(years[i].ToString(), i + 2);
                columnFormats.Add(years[i].ToString(), "#0");
            }
            ExcelWriter resultEW = new ExcelWriter(filePath, "List", resultColumnDic, columnFormats);
            return resultEW;
        }
         
    }
}
