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
    public class ThycGetProvinces : CustomProgramBase
    {
        #region 入口函数 
        public bool Run(string parameters, IListSheet listSheet)
        {
            return GetProvinces(parameters, listSheet);
        }
        #endregion

        #region GetProvinces
        private bool GetProvinces(string parameters, IListSheet listSheet)
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
                "provinceCode"});
             
            string exportDir = this.RunPage.GetExportDir();
             
            string resultFilePath = Path.Combine(exportDir, "途虎养车获取城市列表.xlsx");

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
        /// GetProvinces
        /// </summary>
        /// <param name="listSheet"></param>
        /// <param name="pageSourceDir"></param>
        /// <param name="resultEW"></param>
        private void GetProvinces(IListSheet listSheet, string pageSourceDir, ExcelWriter resultEW)
        {
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                //listSheet中只有一条记录
                string pageUrl = listSheet.PageUrlList[i];
                Dictionary<string, string> row = listSheet.GetRow(i); 
                string localFilePath = this.RunPage.GetFilePath(pageUrl, pageSourceDir);
                HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);

                HtmlNodeCollection allProvinceNodes = htmlDoc.DocumentNode.SelectNodes("//div[@id=\"listTab\"]/ul[1]/li");

                for (int j = 0; j < allProvinceNodes.Count; j++)
                {
                    HtmlNode provinceNode = allProvinceNodes[j];
                    string provinceCode = provinceNode.Attributes["data-value"].Value;
                    string provinceName = provinceNode.InnerText;

                    Dictionary<string, string> f2vs = new Dictionary<string, string>();
                    f2vs.Add("detailPageUrl", "http://www.tuhu.cn/Shops/" + provinceCode + ".aspx");
                    f2vs.Add("detailPageName", provinceCode + provinceName);
                    f2vs.Add("provinceCode", provinceCode);
                    f2vs.Add("provinceName", provinceName); 
                    resultEW.AddRow(f2vs);
                }
            }
        }
        #endregion
    }
}