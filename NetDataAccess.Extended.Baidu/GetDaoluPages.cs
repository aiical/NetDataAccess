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
    public class GetDaoluPages : ExternalRunWebPage
    {
        public override void CheckRequestCompleteFile(string webPageText, Dictionary<string, string> listRow)
        {
            if (webPageText.Contains("current_city") && webPageText.EndsWith("}}"))
            {
            }
            else
            {
                throw new Exception("未完整获取页面");
            }
        }

        public override bool AfterAllGrab(IListSheet listSheet)
        { 
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