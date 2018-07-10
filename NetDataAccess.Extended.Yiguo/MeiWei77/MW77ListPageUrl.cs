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
using NetDataAccess.Base.Server;

namespace NetDataAccess.Extended.Yiguo
{
    /// <summary>
    /// 美味七七
    /// 获取所有列表页地址
    /// </summary>
    public class MW77ListPageUrl : CustomProgramBase
    {
        #region 入口函数
        public bool Run(string parameters, IListSheet listSheet)
        {
            return GetAllListPageUrl(listSheet);
        }
        #endregion

        #region 生成所有列表页地址
        private bool GetAllListPageUrl(IListSheet listSheet)
        {
            string exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            string[] resultColumns = new string[]{"detailPageUrl",
                "detailPageName",
                "cookie",
                "grabStatus", 
                "giveUpGrab", 
                "fullCategoryName",
                "category1Code",
                "category2Code",
                "category3Code",
                "category1Name",
                "category2Name",
                "category3Name",
                "district"};
            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(resultColumns);
            string resultFilePath = Path.Combine(exportDir, this.RunPage.Project.Name + "_AllListPageUrl.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;
            string categoryNameColumnName = SysConfig.DetailPageNameFieldName;

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string fullCategoryName = row[categoryNameColumnName];
                    string url = row[detailPageUrlColumnName];
                    string category1Code = row["category1Code"];
                    string category2Code = row["category2Code"];
                    string category3Code = row["category3Code"];
                    string category1Name = row["category1Name"];
                    string category2Name = row["category2Name"];
                    string category3Name = row["category3Name"];
                    string detailPageUrlPrefix = "http://www.yummy77.com"; 
                    string localFilePath = this.RunPage.GetFilePath(url, pageSourceDir);
                    TextReader tr = null;

                    try
                    {
                        tr = new StreamReader(localFilePath);
                        string webPageHtml = tr.ReadToEnd();

                        HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                        htmlDoc.LoadHtml(webPageHtml);

                        HtmlNode pageNode = htmlDoc.DocumentNode.SelectSingleNode("//div[@class=\"smallpage\"]");
                        int pageCount = 1;
                        if (pageNode != null)
                        {
                            string pageCountStr = pageNode.ChildNodes[2].InnerText.Trim().Split(new string[] { "/" }, StringSplitOptions.RemoveEmptyEntries)[1];
                            pageCount = int.Parse(pageCountStr);
                        } 
                        for (int j = 0; j < pageCount; j++)
                        {
                            string pageIndexStr = (j + 1).ToString();
                            string detailPageName = fullCategoryName + "_" + pageIndexStr;
                            string detailPageUrl = detailPageUrlPrefix + "/category/" + pageIndexStr + "-" + category1Code + "-" + category2Code + "-" + category3Code + "-5-p-.html";
                            Dictionary<string, string> f2vs = new Dictionary<string, string>();
                            f2vs.Add("detailPageUrl", detailPageUrl);
                            f2vs.Add("detailPageName", detailPageName);
                            f2vs.Add("fullCategoryName", fullCategoryName);
                            f2vs.Add("category1Code", category1Code);
                            f2vs.Add("category2Code", category2Code);
                            f2vs.Add("category3Code", category3Code);
                            f2vs.Add("category1Name", category1Name);
                            f2vs.Add("category2Name", category2Name);
                            f2vs.Add("category3Name", category3Name);                  
                            resultEW.AddRow(f2vs);
                        }
                    }
                    catch (Exception ex)
                    {
                        if (tr != null)
                        {
                            tr.Dispose();
                            tr = null;
                        }
                        this.RunPage.InvokeAppendLogText("读取出错.  " + ex.Message + " LocalPath = " + localFilePath, LogLevelType.Error, true);
                        throw ex;
                    }
                }
            }
            resultEW.SaveToDisk();

            //执行后续任务
            TaskManager.StartTask("易果", "美味77获取所有列表页", resultFilePath, null, null, false);
            
            return true;
        }
        #endregion
    }
}