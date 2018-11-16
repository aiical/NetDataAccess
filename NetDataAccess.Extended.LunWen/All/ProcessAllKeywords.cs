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
            try
            {
                string[] parameters = this.Parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                string allKeywordsFilePaht = parameters[0];
                string exportFilePath = parameters[1];


                ExcelReader er = new ExcelReader(allKeywordsFilePaht);
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
                return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
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
