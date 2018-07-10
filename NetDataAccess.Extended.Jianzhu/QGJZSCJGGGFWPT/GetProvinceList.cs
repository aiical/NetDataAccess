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
using System.Web;
using NetDataAccess.Base.Reader;

namespace NetDataAccess.Extended.Jianzhu.QGJZSCJGGGFWPT
{
    public class GetProvinceList : ExternalRunWebPage
    {

        #region GetProvinces
        public override bool AfterAllGrab(IListSheet listSheet)
        { 
            //已经下载下来的首页html保存到的目录（文件夹）
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            
            //输出excel表格包含的列
            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{ 
                "detailPageUrl",
                "detailPageName", 
                "cookie",
                "grabStatus", 
                "giveUpGrab",
                "regionId", 
                "regionName",
                "regionFullName",
                "aptCode",
                "aptScope"});
             
            string exportDir = this.RunPage.GetExportDir();

            string resultFilePath = Path.Combine(exportDir, "企业数据_各省份首页.xlsx");

            //输出对象
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic);

            //从下载到的首页html中，获取列表页个数，并形成所有列表页url
            GetProvinces(listSheet, pageSourceDir, resultEW);

            //保存到硬盘
            resultEW.SaveToDisk();

            return true;
        }
        #endregion

        #region GetProvinces
        /// <summary>
        /// GetCities
        /// </summary>
        /// <param name="listSheet"></param>
        /// <param name="pageSourceDir"></param>
        /// <param name="resultEW"></param>
        private void GetProvinces(IListSheet listSheet, string pageSourceDir, ExcelWriter resultEW)
        {
            string[] paramStrs = this.Parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            string qualListFilePath = paramStrs[0];

            ExcelReader qualER = new ExcelReader(qualListFilePath);
            List<string> qualCodeList = new List<string>();
            int qualCount = qualER.GetRowCount(); 

            for (int i = 0; i < listSheet.RowCount; i++)
            { 
                string pageUrl = listSheet.PageUrlList[i];
                Dictionary<string, string> row = listSheet.GetRow(i); 

                string localFilePath = this.RunPage.GetFilePath(pageUrl, pageSourceDir);
                string fileText = FileHelper.GetTextFromFile(localFilePath);
                 
                JObject rootJo = JObject.Parse(fileText);
                JArray allProvinceObjects = rootJo.SelectToken("json").SelectToken("category").SelectToken("provinces") as JArray;
                for (int j = 0; j < allProvinceObjects.Count; j++)
                {
                    JObject provinceObject = allProvinceObjects[j] as JObject;
                    string regionId = (provinceObject.SelectToken("region_id") as JValue).Value.ToString();
                    string regionName = (provinceObject.SelectToken("region_name") as JValue).Value.ToString();
                    string regionFullName = (provinceObject.SelectToken("region_fullname") as JValue).Value.ToString();
                    for (int k = 0; k < qualCount; k++)
                    {
                        Dictionary<string, string> qual = qualER.GetFieldValues(k);
                        string aptCode = qual["aptCode"];
                        string aptScope = qual["aptScope"];
                        Dictionary<string, string> f2vs = new Dictionary<string, string>();
                        f2vs.Add("detailPageUrl", "http://jzsc.mohurd.gov.cn/dataservice/query/comp/list?regionId=" + regionId + "&aptCode=" + aptCode);
                        f2vs.Add("detailPageName", regionId + aptCode);
                        f2vs.Add("cookie", "filter_comp=show; JSESSIONID=DC4BC03F99DEDEBEFEE739B680BC5230; Hm_lvt_b1b4b9ea61b6f1627192160766a9c55c=1513578016,1513646440,1514281557,1514350446; Hm_lpvt_b1b4b9ea61b6f1627192160766a9c55c=1514356771");
                        f2vs.Add("regionId", regionId);
                        f2vs.Add("regionName", regionName);
                        f2vs.Add("regionFullName", regionFullName);
                        f2vs.Add("aptCode", aptCode);
                        f2vs.Add("aptScope", aptScope);
                        resultEW.AddRow(f2vs);
                    }
                }
            }
        }
        #endregion
    }
}