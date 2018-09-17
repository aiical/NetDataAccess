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
using NetDataAccess.Base.UserAgent;
using System.Net;
using NetDataAccess.Base.Reader;

namespace NetDataAccess.Extended.Dzdp
{
    public class SplitCityStoreInfos : ExternalRunWebPage
    {  
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            string[] parameters = this.Parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            string sourceFilePath = parameters[0];
            string exportDir = parameters[1];
            string cityName = parameters[2];

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("city", 0);
            resultColumnDic.Add("distrctName", 1);
            resultColumnDic.Add("shopName", 2);
            resultColumnDic.Add("shopCode", 3);
            resultColumnDic.Add("address", 4);
            resultColumnDic.Add("tel", 5);
            resultColumnDic.Add("shopType", 6);
            resultColumnDic.Add("commentNum", 7);
            resultColumnDic.Add("lat", 8);
            resultColumnDic.Add("lng", 9);
            resultColumnDic.Add("人均", 10);
            resultColumnDic.Add("口味", 11);
            resultColumnDic.Add("环境", 12);
            resultColumnDic.Add("服务", 13);
            string resultFilePath = Path.Combine(exportDir, "大众点评店铺信息" + cityName + ".xlsx");
            Dictionary<string, string> resultColumnFormat = new Dictionary<string, string>();
            resultColumnFormat.Add("reviewNum", "#,##0");
            resultColumnFormat.Add("lat", "#,##0.000000");
            resultColumnFormat.Add("lng", "#,##0.000000");
            resultColumnFormat.Add("人均", "#,##0.00");
            resultColumnFormat.Add("环境", "#,##0.0");
            resultColumnFormat.Add("口味", "#,##0.0");
            resultColumnFormat.Add("服务", "#,##0.0");

            CsvWriter resultEW = new CsvWriter(resultFilePath, resultColumnDic);

            CsvReader cr = new CsvReader(sourceFilePath);
            int sourceRowCount = cr.GetRowCount();

            for (int i = 0; i < sourceRowCount; i++)
            {
                Dictionary<string, string> sourceRow = cr.GetFieldValues(i);
                string city = sourceRow["city"];
                if (city == cityName)
                {
                    resultEW.AddRow(sourceRow);
                }
            }

            resultEW.SaveToDisk();

            return true;
        } 
    }
}