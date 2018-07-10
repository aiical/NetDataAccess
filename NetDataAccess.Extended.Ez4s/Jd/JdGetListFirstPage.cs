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
    public class JdGetListFirstPage : CustomProgramBase
    {
        #region 入口函数 
        public bool Run(string parameters, IListSheet listSheet)
        {
            return GetListFirstPage(parameters, listSheet);
        }
        #endregion

        #region GetListFirstPage
        private bool GetListFirstPage(string parameters, IListSheet listSheet)
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
                "cat2Name", 
                "pageIndex"});
             
            string exportDir = this.RunPage.GetExportDir();

            string resultFilePath = Path.Combine(exportDir, "京东获取店铺服务所有列表页.xlsx");

            //输出对象
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic);

            //从下载到的首页html中，获取列表页个数，并形成所有列表页url
            GetListFirstPage(listSheet, pageSourceDir, serviceCatFilePath, resultEW);

            //保存到硬盘
            resultEW.SaveToDisk();

            return true;
        }
        #endregion 

        #region GetListFirstPage
        /// <summary>
        /// GetListFirstPage
        /// </summary>
        /// <param name="listSheet"></param>
        /// <param name="pageSourceDir"></param>
        /// <param name="resultEW"></param>
        private void GetListFirstPage(IListSheet listSheet, string pageSourceDir, string serviceCatFilePath, ExcelWriter resultEW)
        {
            for (int i = 0; i < listSheet.RowCount; i++)
            { 
                string pageUrl = listSheet.PageUrlList[i];
                Dictionary<string, string> row = listSheet.GetRow(i);
                string areaLevel1Code = row["areaLevel1Code"];
                string areaLevel1Name = row["areaLevel1Name"];
                string areaLevel2Code = row["areaLevel2Code"];
                string areaLevel2Name = row["areaLevel2Name"];
                string areaLevel3Code = row["areaLevel3Code"];
                string areaLevel3Name = row["areaLevel3Name"];
                string cat1Name = row["cat1Name"];
                string cat2Name = row["cat2Name"];
                string cat2Code = row["cat2Code"];

                string localFilePath = this.RunPage.GetFilePath(pageUrl, pageSourceDir);
                string fileText = FileHelper.GetTextFromFile(localFilePath);

                int jsonBeginIndex = fileText.IndexOf("{");
                int jsonEndIndex = fileText.LastIndexOf("}");

                string jsonStr = fileText.Substring(jsonBeginIndex, jsonEndIndex - jsonBeginIndex + 1);
                JObject rootJo = JObject.Parse(jsonStr);
                JObject dataObject = rootJo.SelectToken("data") as JObject;
                if (dataObject != null)
                {
                    JToken valueToken = dataObject.SelectToken("totalSize");
                    if (valueToken != null)
                    {
                        int totalSize = int.Parse((valueToken as JValue).Value.ToString());

                        for (int j = 0; j < totalSize; j++)
                        {
                            int pageIndex = j + 1;
                            string fullAreaCode = areaLevel1Code + "-" + areaLevel2Code + (areaLevel3Code.Length == 0 ? "" : ("-" + areaLevel3Code));
                            Dictionary<string, object> f2vs = new Dictionary<string, object>();
                            f2vs.Add("detailPageUrl", "http://autobeta.jd.com/service/queryShopAndSku?areaIds=" + fullAreaCode + "&catId=" + cat2Code + "&pageIndex=" + pageIndex.ToString() + "&callback=jQuery3577987&_=1469733824034");
                            f2vs.Add("detailPageName", areaLevel3Code + areaLevel3Name + areaLevel2Code + areaLevel2Name + "_" + cat2Code + "_" + pageIndex.ToString());
                            f2vs.Add("areaLevel1Code", areaLevel1Code);
                            f2vs.Add("areaLevel1Name", areaLevel1Name);
                            f2vs.Add("areaLevel2Code", areaLevel2Code);
                            f2vs.Add("areaLevel2Name", areaLevel2Name);
                            f2vs.Add("areaLevel3Code", areaLevel3Code);
                            f2vs.Add("areaLevel3Name", areaLevel3Name);
                            f2vs.Add("cat1Name", cat1Name);
                            f2vs.Add("cat2Code", cat2Code);
                            f2vs.Add("cat2Name", cat2Name);
                            f2vs.Add("pageIndex", pageIndex);
                            resultEW.AddRow(f2vs);
                        }
                    }
                }
            }
        }
        #endregion
    }
}