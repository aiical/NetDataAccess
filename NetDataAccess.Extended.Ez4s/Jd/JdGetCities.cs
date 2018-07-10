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

namespace NetDataAccess.Extended.Ez4s
{
    public class JdGetCities : CustomProgramBase
    {
        #region 入口函数 
        public bool Run(string parameters, IListSheet listSheet)
        {
            return GetCities(parameters, listSheet);
        }
        #endregion

        #region GetCities
        private bool GetCities(string parameters, IListSheet listSheet)
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
                "areaLevel1Code", 
                "areaLevel1Name",
                "areaLevel2Code", 
                "areaLevel2Name"});
             
            string exportDir = this.RunPage.GetExportDir();
             
            string resultFilePath = Path.Combine(exportDir, "京东服务获取区.xlsx");

            //输出对象
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic);

            //从下载到的首页html中，获取列表页个数，并形成所有列表页url
            GetCities(listSheet, pageSourceDir, resultEW);

            //保存到硬盘
            resultEW.SaveToDisk();

            return true;
        }
        #endregion

        #region GetCities
        /// <summary>
        /// GetCities
        /// </summary>
        /// <param name="listSheet"></param>
        /// <param name="pageSourceDir"></param>
        /// <param name="resultEW"></param>
        private void GetCities(IListSheet listSheet, string pageSourceDir, ExcelWriter resultEW)
        {
            for (int i = 0; i < listSheet.RowCount; i++)
            { 
                string pageUrl = listSheet.PageUrlList[i];
                Dictionary<string, string> row = listSheet.GetRow(i);
                string areaLevel1Code = row["areaLevel1Code"];
                string areaLevel1Name = row["areaLevel1Name"];

                string localFilePath = this.RunPage.GetFilePath(pageUrl, pageSourceDir);
                string fileText = FileHelper.GetTextFromFile(localFilePath);

                int jsonBeginIndex = fileText.IndexOf("{");
                int jsonEndIndex = fileText.LastIndexOf("}");

                string jsonStr = fileText.Substring(jsonBeginIndex, jsonEndIndex - jsonBeginIndex + 1);
                JObject rootJo = JObject.Parse(jsonStr);
                JArray allAreaObjects = rootJo.SelectToken("data") as JArray;
                for (int j = 0; j < allAreaObjects.Count; j++)
                {
                    JObject areaObject = allAreaObjects[j] as JObject;
                    string areaLevel2Code = (areaObject.SelectToken("id") as JValue).Value.ToString();
                    string areaLevel2Name = (areaObject.SelectToken("name") as JValue).Value.ToString();
                    Dictionary<string, string> f2vs = new Dictionary<string, string>();
                    f2vs.Add("detailPageUrl", "http://autobeta.jd.com/queryAreaList?area_lev=3&area_id=" + areaLevel2Code + "&callback=jQuery7711772&_=1469734421125");
                    f2vs.Add("detailPageName", areaLevel2Code + areaLevel2Name);
                    f2vs.Add("areaLevel1Code", areaLevel1Code);
                    f2vs.Add("areaLevel1Name", areaLevel1Name);
                    f2vs.Add("areaLevel2Code", areaLevel2Code);
                    f2vs.Add("areaLevel2Name", areaLevel2Name); 
                    resultEW.AddRow(f2vs);
                }
            }
        }
        #endregion
    }
}