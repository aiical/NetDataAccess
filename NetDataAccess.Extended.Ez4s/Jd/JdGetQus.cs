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

namespace NetDataAccess.Extended.Ez4s
{
    public class JdGetQus : CustomProgramBase
    {
        #region 入口函数 
        public bool Run(string parameters, IListSheet listSheet)
        {
            return GetQus(parameters, listSheet);
        }
        #endregion

        #region GetQus
        private bool GetQus(string parameters, IListSheet listSheet)
        {
            string serviceCatFilePath = parameters;

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
                "areaLevel2Name",
                "areaLevel3Code", 
                "areaLevel3Name",
                "cat1Name", 
                "cat2Code", 
                "cat2Name"});
             
            string exportDir = this.RunPage.GetExportDir();

            string resultFilePath = Path.Combine(exportDir, "京东获取店铺服务首页.xlsx");

            //输出对象
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic);

            //从下载到的首页html中，获取列表页个数，并形成所有列表页url
            GetQus(listSheet, pageSourceDir, serviceCatFilePath, resultEW);

            //保存到硬盘
            resultEW.SaveToDisk();

            return true;
        }
        #endregion 

        #region GetQus
        /// <summary>
        /// GetQus
        /// </summary>
        /// <param name="listSheet"></param>
        /// <param name="pageSourceDir"></param>
        /// <param name="resultEW"></param>
        private void GetQus(IListSheet listSheet, string pageSourceDir, string serviceCatFilePath, ExcelWriter resultEW)
        {

            ExcelReader serviceCatReader = new ExcelReader(serviceCatFilePath);
            int cat2Count = serviceCatReader.GetRowCount();


            for (int i = 0; i < listSheet.RowCount; i++)
            { 
                string pageUrl = listSheet.PageUrlList[i];
                Dictionary<string, string> row = listSheet.GetRow(i);
                string areaLevel1Code = row["areaLevel1Code"];
                string areaLevel1Name = row["areaLevel1Name"];
                string areaLevel2Code = row["areaLevel2Code"];
                string areaLevel2Name = row["areaLevel2Name"];

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
                    string areaLevel3Code = (areaObject.SelectToken("id") as JValue).Value.ToString();
                    string areaLevel3Name = (areaObject.SelectToken("name") as JValue).Value.ToString();
                    for (int k = 0; k < cat2Count; k++)
                    {
                        Dictionary<string, string> catRow = serviceCatReader.GetFieldValues(k);
                        string cat1Name = catRow["cat1Name"];
                        string cat2Code = catRow["cat2Code"];
                        string cat2Name = catRow["cat2Name"];
                        string fullAreaCode = areaLevel1Code + "-" + areaLevel2Code + "-" + areaLevel3Code;
                        Dictionary<string, string> f2vs = new Dictionary<string, string>();
                        f2vs.Add("detailPageUrl", "http://autobeta.jd.com/service/queryShopAndSku?areaIds=" + fullAreaCode + "&catId=" + cat2Code + "&pageIndex=1&callback=jQuery3577987&_=1469733824034");
                        f2vs.Add("detailPageName", areaLevel3Code + areaLevel3Name + cat2Code);
                        f2vs.Add("areaLevel1Code", areaLevel1Code);
                        f2vs.Add("areaLevel1Name", areaLevel1Name);
                        f2vs.Add("areaLevel2Code", areaLevel2Code);
                        f2vs.Add("areaLevel2Name", areaLevel2Name);
                        f2vs.Add("areaLevel3Code", areaLevel3Code);
                        f2vs.Add("areaLevel3Name", areaLevel3Name);
                        f2vs.Add("cat1Name", cat1Name);
                        f2vs.Add("cat2Code", cat2Code);
                        f2vs.Add("cat2Name", cat2Name);
                        resultEW.AddRow(f2vs);
                    }
                }
                for (int k = 0; k < cat2Count; k++)
                {
                    Dictionary<string, string> catRow = serviceCatReader.GetFieldValues(k);
                    string cat1Name = catRow["cat1Name"];
                    string cat2Code = catRow["cat2Code"];
                    string cat2Name = catRow["cat2Name"];
                    string fullAreaCode = areaLevel1Code + "-" + areaLevel2Code;
                    Dictionary<string, string> f2vs = new Dictionary<string, string>();
                    f2vs.Add("detailPageUrl", "http://autobeta.jd.com/service/queryShopAndSku?areaIds=" + fullAreaCode + "&catId=" + cat2Code + "&pageIndex=1&callback=jQuery3577987&_=1469733824034");
                    f2vs.Add("detailPageName", areaLevel2Code + areaLevel2Name + cat2Code);
                    f2vs.Add("areaLevel1Code", areaLevel1Code);
                    f2vs.Add("areaLevel1Name", areaLevel1Name);
                    f2vs.Add("areaLevel2Code", areaLevel2Code);
                    f2vs.Add("areaLevel2Name", areaLevel2Name);
                    f2vs.Add("areaLevel3Code", "");
                    f2vs.Add("areaLevel3Name", "");
                    f2vs.Add("cat1Name", cat1Name);
                    f2vs.Add("cat2Code", cat2Code);
                    f2vs.Add("cat2Name", cat2Name);
                    resultEW.AddRow(f2vs);
                }
            }
        }
        #endregion
    }
}