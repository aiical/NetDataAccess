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
    public class GetUserSearchPageUrls : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            string[] paramters = this.Parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            string nameListFilePath = paramters[0];
            string outputFilePath = paramters[1];

            ExcelReader er = new ExcelReader(nameListFilePath, "List");
            int rowCount = er.GetRowCount();

            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                "detailPageUrl",
                "detailPageName", 
                "cookie",
                "grabStatus", 
                "giveUpGrab", 
                "FirmID",
                "FirmName",
                "LastName",
                "FirstName",
                "MiddleName"});
            ExcelWriter resultEW = new ExcelWriter(outputFilePath, "List", resultColumnDic);

            for (int i = 0; i < rowCount; i++)
            {
                Dictionary<string, string> row = er.GetFieldValues(i); 
                string lastName = row["LastName"];
                string firstName = row["FirstName"];
                string nameUrl = CommonUtil.UrlEncode(firstName + " " + lastName);
                string pageUrl = "https://cn.bing.com/search?q=" + nameUrl + "%20site%3Awww.linkedin.com%2Fin&qs=n&form=QBRE&sp=-1&pq=" + nameUrl + "%20site%3Awww.linkedin.com%2Fin&sc=0-40&sk=&cvid=8BFDE45D86074A869AF2964EF24A0127&ttt=" + i.ToString();
                row["detailPageUrl"] = pageUrl;
                row["detailPageName"] = pageUrl;

                resultEW.AddRow(row);
            }
            resultEW.SaveToDisk();
            
            return base.AfterAllGrab(listSheet);
        }
    }
}