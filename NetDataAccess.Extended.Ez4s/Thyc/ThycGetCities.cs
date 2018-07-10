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
    public class ThycGetCities : CustomProgramBase
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
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            
            //输出excel表格包含的列
            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                "detailPageUrl",
                "detailPageName", 
                "cookie",
                "grabStatus", 
                "giveUpGrab",
                "provinceName",
                "provinceCode",
                "cityName",
                "cityCode"});
             
            string exportDir = this.RunPage.GetExportDir();
             
            string resultFilePath = Path.Combine(exportDir, "途虎养车获取维修站列表.xlsx");

            //输出对象
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic);
             
            this.ReadCityPages(listSheet, pageSourceDir, resultEW);

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
        private void ReadCityPages(IListSheet listSheet, string pageSourceDir, ExcelWriter resultEW)
        {
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                //listSheet中只有一条记录
                Dictionary<string, string> row = listSheet.GetRow(i);
                string pageUrl = row[SysConfig.DetailPageUrlFieldName];
                string provinceCode = row["provinceCode"];
                string provinceName = row["provinceName"];
                string localFilePath = this.RunPage.GetFilePath(pageUrl, pageSourceDir); 
                HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);

                HtmlNodeCollection allCityNodes = htmlDoc.DocumentNode.SelectNodes("//div[@id=\"listTab\"]/ul[2]/li/a");

                for (int j = 0; j < allCityNodes.Count; j++)
                {
                    HtmlNode cityNode = allCityNodes[j];
                    string cityUrl = cityNode.Attributes["href"].Value;
                    string[] cityUrlPieces = cityUrl.Split(new string[] { "/" }, StringSplitOptions.RemoveEmptyEntries);
                    string[] cityPageNamePieces = cityUrlPieces[cityUrlPieces.Length - 1].Split(new string[] { "." }, StringSplitOptions.RemoveEmptyEntries);
                    string cityCode = cityPageNamePieces[0];
                    string cityName = cityNode.InnerText;

                    Dictionary<string, string> f2vs = new Dictionary<string, string>();
                    f2vs.Add("detailPageUrl", cityUrl);
                    f2vs.Add("detailPageName", cityCode + cityName);
                    f2vs.Add("provinceCode", provinceCode);
                    f2vs.Add("provinceName", provinceName);
                    f2vs.Add("cityCode", cityCode);
                    f2vs.Add("cityName", cityName);
                    resultEW.AddRow(f2vs);
                }
            }
        }
        #endregion
    }
}