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
    public class JdGetCats : CustomProgramBase
    {
        #region 入口函数 
        public bool Run(string parameters, IListSheet listSheet)
        {
            return  GetCats(parameters, listSheet);
        }
        #endregion

        #region GetCats
        private bool GetCats(string parameters, IListSheet listSheet)
        { 
            //已经下载下来的首页html保存到的目录（文件夹）
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            
            //输出excel表格包含的列
            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{ 
                "cat1Name", 
                "cat2Name",
                "cat2Code"});
             
            string exportDir = this.RunPage.GetExportDir();
             
            string resultFilePath = Path.Combine(exportDir, "京东服务分类.xlsx");

            //输出对象
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic);

            //从下载到的首页html中，获取列表页个数，并形成所有列表页url
            GetCats(listSheet, pageSourceDir, resultEW);

            //保存到硬盘
            resultEW.SaveToDisk();

            return true;
        }
        #endregion

        #region GetCats
        /// <summary>
        /// GetCats
        /// </summary>
        /// <param name="listSheet"></param>
        /// <param name="pageSourceDir"></param>
        /// <param name="resultEW"></param>
        private void GetCats(IListSheet listSheet, string pageSourceDir, ExcelWriter resultEW)
        {
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                //listSheet中只有一条记录
                string pageUrl = listSheet.PageUrlList[i];
                Dictionary<string, string> row = listSheet.GetRow(i); 
                string localFilePath = this.RunPage.GetFilePath(pageUrl, pageSourceDir);
                HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);

                HtmlNodeCollection allCat1Nodes = htmlDoc.DocumentNode.SelectNodes("//div[@id=\"catDiv\"]/div/h5");
                HtmlNodeCollection allCat2GroupNodes = htmlDoc.DocumentNode.SelectNodes("//div[@id=\"catDiv\"]/div/ul");

                for (int j = 0; j < allCat1Nodes.Count; j++)
                {
                    HtmlNode cat1Node = allCat1Nodes[j];
                    HtmlNode cat2GroupNode = allCat2GroupNodes[j];
                    string cat1Name = cat1Node.InnerText.Trim();
                    HtmlNodeCollection allCat2Nodes = cat2GroupNode.SelectNodes("./li");
                    for (int k = 0; k < allCat2Nodes.Count; k++)
                    {
                        HtmlNode cat2Node = allCat2Nodes[k];
                        string cat2Code = cat2Node.Attributes["catid"].Value;
                        string cat2Name = cat2Node.InnerText.Trim();

                        Dictionary<string, string> f2vs = new Dictionary<string, string>(); 
                        f2vs.Add("cat1Name", cat1Name);
                        f2vs.Add("cat2Name", cat2Name);
                        f2vs.Add("cat2Code", cat2Code); 
                        resultEW.AddRow(f2vs);
                    }
                }
            }
        }
        #endregion
    }
}