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

namespace NetDataAccess.Extended.Yiguo
{
    public class WomaiListPageUrl : CustomProgramBase
    {
        public bool Run(string parameters, IListSheet listSheet)
        {
            return GetAllListPageUrl(listSheet);
        }
        private bool GetAllListPageUrl(IListSheet listSheet)
        {
            string exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            
            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("fullCategoryName", 5);
            resultColumnDic.Add("category1Code", 6);
            resultColumnDic.Add("category2Code", 7); 
            resultColumnDic.Add("category1Name", 8);
            resultColumnDic.Add("category2Name", 9); 
            resultColumnDic.Add("district", 10);
            string resultFilePath = Path.Combine(exportDir, "我买网获取所有列表页.xlsx");
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
                    string cookie = row["cookie"];
                    string category1Code = row["category1Code"];
                    string category2Code = row["category2Code"]; 
                    string category1Name = row["category1Name"];
                    string category2Name = row["category2Name"]; 
                    string district = row["district"]; 
                    string localFilePath = this.RunPage.GetFilePath(url, pageSourceDir);
                    TextReader tr = null;

                    try
                    {
                        tr = new StreamReader(localFilePath);
                        string webPageHtml = tr.ReadToEnd();

                        HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                        htmlDoc.LoadHtml(webPageHtml);

                        HtmlNode pageNode = htmlDoc.DocumentNode.SelectSingleNode("//div[@class=\"result-sum clearfix\"]");
                        if (pageNode != null)
                        {
                            string pageStr = pageNode.InnerText.Trim();
                            string[] pageSplits = pageStr.Split(new string[] { "/" }, StringSplitOptions.RemoveEmptyEntries);
                            string pageCountStr = pageSplits[1];
                            int pageCount = int.Parse(pageCountStr);
                            for (int j = 0; j < pageCount; j++)
                            {
                                string pageIndexStr = (j + 1).ToString();
                                string detailPageName = fullCategoryName + "_" + pageIndexStr;
                                string detailPageUrl = "http://www.womai.com/ProductList.do?mid=0&Cid=" + category2Code + "&mainColumnId=&page=" + pageIndexStr + "&brand=-1&rypId=608&zhId=605&orderBy=&isPromotions=&sellable=&Keywords=&Keyword=&isKeyCommendClick=1&sellerid=&selAttr=&selCol=&urllist=";
                                Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                f2vs.Add("detailPageUrl", detailPageUrl);
                                f2vs.Add("detailPageName", detailPageName);
                                f2vs.Add("cookie", cookie);
                                f2vs.Add("fullCategoryName", fullCategoryName);
                                f2vs.Add("category1Code", category1Code);
                                f2vs.Add("category2Code", category2Code);
                                f2vs.Add("category1Name", category1Name);
                                f2vs.Add("category2Name", category2Name);
                                f2vs.Add("district", district);
                                resultEW.AddRow(f2vs);
                            }
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

            return true;
        }
    }
}