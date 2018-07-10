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
    public class YcwyGetCities : CustomProgramBase
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
                "provinceName", 
                "cityName",
                "cityCode"});
             
            string exportDir = this.RunPage.GetExportDir();
             
            string resultFilePath = Path.Combine(exportDir, "养车无忧列表页.xlsx");

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
                //listSheet中只有一条记录
                string pageUrl = listSheet.PageUrlList[i];
                Dictionary<string, string> row = listSheet.GetRow(i); 
                string localFilePath = this.RunPage.GetFilePath(pageUrl, pageSourceDir);
                HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);

                HtmlNodeCollection allProvinceNodes = htmlDoc.DocumentNode.SelectNodes("//div[@class=\"province\"]/b");
                HtmlNodeCollection allCityListNodes = htmlDoc.DocumentNode.SelectNodes("//div[@class=\"province\"]/p");

                for (int j = 0; j < allProvinceNodes.Count; j++)
                {
                    HtmlNode provinceNode = allProvinceNodes[j];
                    HtmlNode cityListNode = allCityListNodes[j];
                    string provinceName = provinceNode.InnerText.Trim().Replace("：", "");
                    HtmlNodeCollection allCityNodes =  cityListNode.SelectNodes("./a");
                    for (int k = 0; k < allCityNodes.Count; k++)
                    {
                        HtmlNode cityNode = allCityNodes[k];
                        string cityUrl = cityNode.Attributes["href"].Value;
                        string[] cityPieces = cityUrl.Split(new string[] { "/" }, StringSplitOptions.RemoveEmptyEntries);
                        string cityCode = cityPieces[cityPieces.Length - 1];
                        string cityName = cityNode.InnerText.Trim();

                        Dictionary<string, string> f2vs = new Dictionary<string, string>();
                        f2vs.Add("detailPageUrl", cityUrl);
                        f2vs.Add("detailPageName", cityCode + cityName);
                        f2vs.Add("provinceName", provinceName);
                        f2vs.Add("cityCode", cityCode);
                        f2vs.Add("cityName", cityName); 
                        resultEW.AddRow(f2vs);
                    }
                }
            }
        }
        #endregion
    }
}