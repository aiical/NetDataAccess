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

namespace NetDataAccess.Extended.Keywords.WanFang
{
    public class GetKeywordPercents : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            try
            {
                string[] parameters = this.Parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                string keywordListFilePath = parameters[0];
                string resultDir= parameters[1];
                string categoryName= parameters[2];
                string[] subCategoryNames = parameters[3].Split(new string[] { "|" }, StringSplitOptions.RemoveEmptyEntries);

                Dictionary<string, double> keywordsSumValueDic = new Dictionary<string, double>(); 
                Dictionary<string, Dictionary<string, double>> categoryKeywordsValue = new Dictionary<string, Dictionary<string, double>>();
                foreach (string subCategoryName in subCategoryNames)
                {
                    Dictionary<string, double> keywordsNum = new Dictionary<string, double>();
                    categoryKeywordsValue.Add(subCategoryName, keywordsNum);

                    ExcelReader er = new ExcelReader(keywordListFilePath, subCategoryName);
                    int rowCount = er.GetRowCount();
                    int keywordCount = 0;
                    for (int i = 0; i < rowCount; i++)
                    {
                        Dictionary<string, string> keywordRow = er.GetFieldValues(i);
                        string keyword = keywordRow["词汇"];
                        //去掉包含空格的
                        if (!keyword.Contains(" "))
                        {
                            int num = int.Parse(keywordRow["出现次数"]); 

                            keywordCount += num; 
                        }
                    }
                    for (int i = 0; i < rowCount; i++)
                    {
                        Dictionary<string, string> keywordRow = er.GetFieldValues(i);
                        string keyword = keywordRow["词汇"];
                        //去掉包含空格的
                        if (!keyword.Contains(" "))
                        {
                            double value = double.Parse(keywordRow["出现次数"]) / keywordCount;
                            keywordsNum.Add(keyword, value);
                             

                            if (keywordsSumValueDic.ContainsKey(keyword))
                            {
                                keywordsSumValueDic[keyword] = keywordsSumValueDic[keyword] + value;
                            }
                            else
                            {
                                keywordsSumValueDic.Add(keyword, value);
                            }
                        }
                    }
                }

                this.SaveNLPCustomDictionary(resultDir, categoryName, keywordsSumValueDic);

                ExcelWriter resultEW = this.GetSubCategoryPercentsExcelWriter(resultDir, categoryName);
                foreach (string subCategoryName in subCategoryNames)
                {  
                    ExcelReader er = new ExcelReader(keywordListFilePath, subCategoryName);
                    this.GetSubCategoryKeywordPercents(resultEW, categoryName, subCategoryName, er, keywordsSumValueDic, categoryKeywordsValue[subCategoryName]);
                }
                resultEW.SaveToDisk();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return true;
        }

        private void GetSubCategoryKeywordPercents(ExcelWriter resultEW, string categoryName, string subCategoryName, ExcelReader er, Dictionary<string, double> keywordsSumValueDic, Dictionary<string, double> keywordsValueDic)
        {
            int rowCount = er.GetRowCount();
            for (int i = 0; i < rowCount; i++)
            {
                Dictionary<string, string> keywordRow = er.GetFieldValues(i);
                string keyword = keywordRow["词汇"];
                //去掉包含空格的
                if (!keyword.Contains(" "))
                {
                    double value = keywordsValueDic[keyword];
                    double sumValue = keywordsSumValueDic[keyword];
                    double percent = (double)value / (double)sumValue;
                    if (value > 0.0001)
                    {
                        Dictionary<string, object> resultRow = new Dictionary<string, object>();
                        resultRow.Add("category", categoryName);
                        resultRow.Add("subCategory", subCategoryName);
                        resultRow.Add("keyword", keyword);
                        resultRow.Add("percent", percent);
                        resultEW.AddRow(resultRow);
                    }
                }
            }
        } 

        private void SaveNLPCustomDictionary(string resultDir, string categoryName, Dictionary<string, double> keywordsSumValueDic)
        {
            StringBuilder ss = new StringBuilder();
            foreach(string keyword in keywordsSumValueDic.Keys)
            { 
                ss.AppendLine(keyword + "  n 1");
            }
            String resultFilePath = Path.Combine(resultDir, "dic_" + categoryName + ".txt");
            FileHelper.SaveTextToFile(ss.ToString(), resultFilePath);
        }

        private ExcelWriter GetSubCategoryPercentsExcelWriter(string resultDir, string categoryName)
        {
            String resultFilePath = Path.Combine(resultDir, "keywordPercents_" + categoryName + ".xlsx");
            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[] { 
                "category", 
                "subCategory", 
                "keyword", 
                "percent"});
            Dictionary<string, string> columnFormats = new Dictionary<string, string>();
            columnFormats.Add("percent", "#0.00000000");

            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, columnFormats);
            return resultEW;
        } 
    }
}