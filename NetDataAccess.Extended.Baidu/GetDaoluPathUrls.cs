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

namespace NetDataAccess.Extended.Baidu
{
    public class GetDaoluPathUrls : ExternalRunWebPage
    {

        public override bool AfterAllGrab(IListSheet listSheet)
        {
            string[] allParameters = this.Parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            string daoluListFilePath = allParameters[0];
            string exportDir = allParameters[1];

            CsvReader cr = new CsvReader(daoluListFilePath);
            int rowCount = cr.GetRowCount();
            int fileIndex = 1;
            ExcelWriter resultEW = null;
            for (int i = 0; i < rowCount; i++)
            {
                Dictionary<string, string> row = cr.GetFieldValues(i);
                if (i % 500000 == 0)
                {
                    if (resultEW != null)
                    {
                        resultEW.SaveToDisk();
                    }
                    resultEW = this.GetExcelWriter(exportDir, fileIndex);
                    fileIndex++;
                }
                string uid = row["uid"];
                Dictionary<string, string> f2vs = new Dictionary<string, string>();
                f2vs.Add("detailPageUrl", "http://map.baidu.com/?newmap=1&reqflag=pcmap&biz=1&from=webmap&da_par=direct&pcevaname=pc4.1&qt=ext&uid=" + uid + "&l=18&c=2032&tn=B_NORMAL_MAP&nn=0");
                f2vs.Add("detailPageName", uid);
                f2vs.Add("uid", uid);
                f2vs.Add("title", row["title"]);
                f2vs.Add("province", row["province"]);
                f2vs.Add("city", row["city"]);
                f2vs.Add("address", row["address"]);
                f2vs.Add("phoneNumber", row["phoneNumber"]);
                f2vs.Add("postcode", row["postcode"]);
                f2vs.Add("url", row["url"]);
                f2vs.Add("lat", row["lat"]);
                f2vs.Add("lng", row["lng"]);
                resultEW.AddRow(f2vs);
            }

            resultEW.SaveToDisk();

            return true;
        }

        private ExcelWriter GetExcelWriter(string exportDir, int fileIndex)
        {
            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                "detailPageUrl",
                "detailPageName", 
                "cookie",
                "grabStatus", 
                "giveUpGrab", 
                "uid", 
                "title",
                "province", 
                "city",  
                "address", 
                "phoneNumber", 
                "postcode",
                "url",
                "lat",
                "lng"});

            string filePath = Path.Combine(exportDir, "获取道路信息_" + fileIndex.ToString() + ".xlsx");

            ExcelWriter resultEW = new ExcelWriter(filePath, "List", resultColumnDic);
            return resultEW;
        }
    }
}