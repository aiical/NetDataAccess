using NetDataAccess.Base.Config;
using NetDataAccess.Base.Reader;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace NetDataAccess.Extended.Linkedin.Common
{
    public class ProcessGooglePage
    {
        private static List<string> _ChineseWords = null;
        private static List<string> ChineseWords
        {
            get
            {
                if (_ChineseWords == null)
                {
                    List<String> chineseWords = new List<string>();
                    string filePath = Path.Combine(SysConfig.SysFileDir, "User/ChineseWords.xlsx");
                    ExcelReader er = new ExcelReader(filePath);
                    int rowCount = er.GetRowCount();
                    for (int i = 0; i < rowCount; i++)
                    {
                        chineseWords.Add(er.GetFieldValues(i)["word"]);
                    }
                    _ChineseWords = chineseWords;
                    er.Close();
                }
                return _ChineseWords;
            }
        }

        private static List<string> _CommonSites = null;
        private static List<string> CommonSites
        {
            get
            {
                if (_CommonSites == null)
                {
                    List<String> commonSites = new List<string>();
                    string filePath = Path.Combine(SysConfig.SysFileDir, "User/CommonSites.xlsx");
                    ExcelReader er = new ExcelReader(filePath);
                    int rowCount = er.GetRowCount();
                    for (int i = 0; i < rowCount; i++)
                    {
                        commonSites.Add(er.GetFieldValues(i)["site"]);
                    }
                    _CommonSites = commonSites;
                    er.Close();
                }
                return _CommonSites;
            }
        }

        public static string GetRandomSearchValue()
        {
            Random random = new Random(DateTime.Now.Millisecond);
            int wordIndex = random.Next(0, ChineseWords.Count);
            int siteIndex = random.Next(0, CommonSites.Count + 50);
            return ChineseWords[wordIndex] + (siteIndex < CommonSites.Count ? (" inurl:" + CommonSites[siteIndex]) : "");
        }
    }
}
