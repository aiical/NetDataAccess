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
using NetDataAccess.Base.CsvHelper;
using NetDataAccess.Base.Reader;

namespace NetDataAccess.Extended.GuPiao
{
    public class SearchGongGao : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            this.Search();
            return true;
        }

        private void Search()
        {
            string[] parameters = this.Parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            string sourceFilePath = parameters[0];
            string destFilePath = parameters[1];
            string keywordsGroupStr = parameters[2];

            string[] keywordsGroup = keywordsGroupStr.Split(new string[] { "$" }, StringSplitOptions.RemoveEmptyEntries);
            List<String[]> keywordsList = new List<string[]>();
            foreach (string keywordsStr in keywordsGroup)
            {
                keywordsList.Add(keywordsStr.Split(new string[] { "|" }, StringSplitOptions.RemoveEmptyEntries));
            }
              
            CsvWriter cw = this.GetCsvWriter(destFilePath);

            CsvReader cr = new CsvReader(sourceFilePath);
            int rowCount = cr.GetRowCount();
            for (int i = 0; i < rowCount; i++)
            {
                Dictionary<string,string> row = cr.GetFieldValues(i);
                string announcementTitle =  row["announcementTitle"];

                for (int j = 0; j < keywordsList.Count; j++)
                {
                    bool matched = true;
                    string[] keywords = keywordsList[j];
                    foreach (string keyword in keywords)
                    {
                        if (!announcementTitle.Contains(keyword))
                        {
                            matched = false;
                            break;
                        }
                    }
                    if (matched)
                    {
                        cw.AddRow(row);
                        break;
                    }
                }
            }
            cw.SaveToDisk();
        }
        private CsvWriter GetCsvWriter(string destFilePath)
        { 

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("pinyin", 0);
            resultColumnDic.Add("zwjc", 1);
            resultColumnDic.Add("code", 2);
            resultColumnDic.Add("orgId", 3);
            resultColumnDic.Add("stockExchange", 4);
            resultColumnDic.Add("category", 5);
            resultColumnDic.Add("announcementTitle", 6);
            resultColumnDic.Add("announcementTime", 7);
            resultColumnDic.Add("announcementId", 8);
            resultColumnDic.Add("adjunctType", 9);
            resultColumnDic.Add("fileUrl", 10); 
            CsvWriter resultEW = new CsvWriter(destFilePath, resultColumnDic);
            return resultEW;
        }         

    }
}