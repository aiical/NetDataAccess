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
using System.Text.RegularExpressions;

namespace NetDataAccess.Extended.Keywords.WanFang
{
    public class GetKeywordTraining : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            try
            {
                string[] parameters = this.Parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                string keywordPercentFilePath = parameters[0];
                string trainingDir = parameters[1];
                string resultDir = parameters[2];
                string categoryName= parameters[3];
                string[] subCategoryNames = parameters[4].Split(new string[] { "|" }, StringSplitOptions.RemoveEmptyEntries);

                Dictionary<string, double> keywordsSumValueDic = new Dictionary<string, double>();
                Dictionary<string, Dictionary<string, double>> categoryKeywordsValue = new Dictionary<string, Dictionary<string, double>>();
                ExcelReader er = new ExcelReader(keywordPercentFilePath);
                ExcelWriter resultEW = this.GetSubCategoryTrainingPercentsExcelWriter(resultDir, categoryName);
                foreach (string subCategoryName in subCategoryNames)
                {
                    this.Training(trainingDir, resultDir, subCategoryName, er, resultEW);
                }
                resultEW.SaveToDisk();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return true;
        }

        private void Training(string trainingFileDir, string resultDir, string subCategoryName, ExcelReader er, ExcelWriter resultEW)
        {
            Dictionary<string, int> wordNumDic = new Dictionary<string, int>();
            string dirPath = Path.Combine(trainingFileDir, subCategoryName);
            string[] filePaths = Directory.GetFiles(dirPath);
            int totalCount = 0;
            foreach (string filePath in filePaths)
            {
                string fileText = FileHelper.GetTextFromFile(filePath);
                for (int i = 0; i < er.GetRowCount(); i++)
                {
                    Dictionary<string, string> sourceRow = er.GetFieldValues(i);
                    string word = sourceRow["keyword"];
                    string cateName = sourceRow.ContainsKey("category") ? sourceRow["category"] : "";
                    string subCateName = sourceRow.ContainsKey("subCategory") ? sourceRow["subCategory"] : "";
                    if (subCateName == subCategoryName)
                    {
                        string kw = this.EscapeExprSpecialWord(word);
                        Regex r = new Regex(kw);
                        MatchCollection mc = r.Matches(fileText);
                        if (mc != null && mc.Count > 0)
                        {
                            if (wordNumDic.ContainsKey(word))
                            {
                                wordNumDic[word] = wordNumDic[word] + mc.Count;
                            }
                            else
                            {
                                wordNumDic.Add(word, mc.Count);
                            }
                            totalCount += mc.Count;
                        }
                    }
                }
            }

            for (int i = 0; i < er.GetRowCount(); i++)
            {
                Dictionary<string, string> sourceRow = er.GetFieldValues(i);
                string word = sourceRow["keyword"];
                string cateName = sourceRow.ContainsKey("category") ? sourceRow["category"] : "";
                string subCateName = sourceRow.ContainsKey("subCategory") ? sourceRow["subCategory"] : "";
                double percent = double.Parse(sourceRow["percent"]);
                if (subCategoryName == subCateName)
                {
                    if (wordNumDic.ContainsKey(word))
                    {
                        double p = wordNumDic[word];
                        if (p > 0)
                        {
                            Dictionary<string, object> resultRow = new Dictionary<string, object>();
                            resultRow.Add("category", cateName);
                            resultRow.Add("subCategory", subCateName);
                            resultRow.Add("keyword", word);
                            resultRow.Add("percent", !wordNumDic.ContainsKey(word) ? 0 : p / (double)totalCount * percent);
                            resultEW.AddRow(resultRow);
                        }
                    }
                }
            }
        }
        private string EscapeExprSpecialWord(String keyword)
        {
            String[] fbsArr = { "\\", "$", "(", ")", "*", "+", ".", "[", "]", "?", "^", "{", "}", "|" };
            foreach (String key in fbsArr)
            {
                if (keyword.Contains(key))
                {
                    keyword = keyword.Replace(key, "\\" + key);
                }
            }
            return keyword;
        }

        private ExcelWriter GetSubCategoryTrainingPercentsExcelWriter(string resultDir, string categoryName)
        {
            String resultFilePath = Path.Combine(resultDir, "keywordTrainingPercents_" + categoryName + ".xlsx");
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