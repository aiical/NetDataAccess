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
using NetDataAccess.Base.Writer;
using NetDataAccess.Base.Reader;
using NetDataAccess.Base.CsvHelper;
using NetDataAccess.Base.DB;
using HtmlAgilityPack;
using System.Web; 
using System.Collections;

namespace NetDataAccess.Extended.Linkedin.KeyWordsSearchAtBing
{
    public class GetKeywordsSearchPageUrls : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            string[] paramters = this.Parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            string nameListFilePath = paramters[0];
            string outputFilePath = paramters[1];
            int searchPageCount = int.Parse(paramters[2]);

            ExcelReader er = new ExcelReader(nameListFilePath, "List");
            int rowCount = er.GetRowCount();

            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                "detailPageUrl",
                "detailPageName", 
                "cookie",
                "grabStatus", 
                "giveUpGrab", 
                "keywords",
                "pageIndex"});
            ExcelWriter resultEW = new ExcelWriter(outputFilePath, "List", resultColumnDic);

            for (int i = 0; i < rowCount; i++)
            {
                for (int j = 0; j < searchPageCount; j++)
                {
                    Dictionary<string, string> row = er.GetFieldValues(i);
                    string keywords = row["keywords"];
                    string keywordsUrl = CommonUtil.UrlEncode(keywords);
                    int beginPageIndex = j * 14 + 1;

                    string pageUrl = "https://cn.bing.com/search?q=" + keywordsUrl + "%20site%3awww.linkedin.com%2fin&qs=n&sp=-1&pq=" + keywordsUrl + "%20site%3awww.linkedin.com%2fin&sc=0-41&sk=&cvid=0EA188BEA7174D4CBD57CDBFDA73B06C&&first=" + beginPageIndex.ToString() + "&FORM=PERE";
                    row["detailPageUrl"] = pageUrl;
                    row["detailPageName"] = pageUrl;

                    row["keywords"] = keywords;

                    row["pageIndex"] = (j + 1).ToString();

                    resultEW.AddRow(row);
                }
            }
            resultEW.SaveToDisk();
            
            return base.AfterAllGrab(listSheet);
        }
    }
}