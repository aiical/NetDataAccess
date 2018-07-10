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

namespace NetDataAccess.Extended.Celloud
{
    public class OalibInitFirstPageUrl : CustomProgramBase
    {
        public bool Run(string parameters, IListSheet listSheet)
        {
            return GetInitAllFirstPageUrl(parameters, listSheet);
        }
        private bool GetInitAllFirstPageUrl(string parameters, IListSheet listSheet)
        {
            string[] parameterArray = parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);

            string exportDir = parameterArray[0];
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            
            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("species", 5);  
            resultColumnDic.Add("year", 6);

            String listFileDir = Path.GetDirectoryName(this.RunPage.ListFilePath);

            string resultFilePath = Path.Combine(listFileDir, "oalib获取列表页首页.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName; 

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                { 
                    string url = row[detailPageUrlColumnName];
                    string cookie = row["cookie"];
                    string species = row["species"].Trim(); 
                    string searchKeyStr = CommonUtil.UrlEncode(species);

                    try
                    {
                        for (int year = 1980; year <= 2016; year++)
                        {
                            string detailPageName = species + "_" + year;
                            string detailPageUrl = "http://www.oalib.com/search?type=0&oldType=0&kw=%22" + searchKeyStr + "%22&searchField=All&__multiselect_searchField=&fromYear=" + year.ToString() + "&toYear=&pageNo=1";
                            Dictionary<string, string> f2vs = new Dictionary<string, string>();
                            f2vs.Add("detailPageUrl", detailPageUrl);
                            f2vs.Add("detailPageName", detailPageName);
                            f2vs.Add("cookie", cookie);
                            f2vs.Add("species", species); 
                            f2vs.Add("year", year.ToString());
                            resultEW.AddRow(f2vs);
                        }
                    }
                    catch (Exception ex)
                    {
                        this.RunPage.InvokeAppendLogText("扩展程序出错.  " + ex.Message, LogLevelType.Error, true);
                        throw ex;
                    }
                }
            }
            resultEW.SaveToDisk(); 

            return true;
        }
    }
}