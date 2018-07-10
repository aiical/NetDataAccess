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
    public class FeiniuDetailPageUrl : CustomProgramBase
    {
        public bool Run(string parameters, IListSheet listSheet)
        {
            return GetAllDetailPageUrl(listSheet);
        }
        private bool GetAllDetailPageUrl(IListSheet listSheet)
        {
            string exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                "detailPageUrl",
                "detailPageName", 
                "cookie",
                "grabStatus", 
                "giveUpGrab", 
                "productCode",
                "productName", 
                "category1Code", 
                "category2Code", 
                "category3Code", 
                "category1Name",
                "category2Name", 
                "category3Name" });
            string resultFilePath = Path.Combine(exportDir, this.RunPage.Project.Name + "_AllDetailPageUrl.xlsx");
            
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic);

            Dictionary<string, string> goodsDic = new Dictionary<string, string>();

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;
            string categoryNameColumnName = SysConfig.DetailPageNameFieldName;

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                { 
                    string url = row[detailPageUrlColumnName]; 
                    string category1Code = row["category1Code"];
                    string category2Code = row["category2Code"];
                    string category3Code = row["category3Code"];
                    string category1Name = row["category1Name"];
                    string category2Name = row["category2Name"];
                    string category3Name = row["category3Name"];   
                    string localFilePath = this.RunPage.GetFilePath(url, pageSourceDir); 

                    try
                    {
                        {
                            HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                            HtmlNodeCollection itemNodes = htmlDoc.DocumentNode.SelectNodes("//ul[@id=\"cata_choose_product\"]/li");
                            if (itemNodes != null)
                            {
                                foreach (HtmlNode itemNode in itemNodes)
                                {
                                    //HtmlNodeCollection allPageNodes = listNode.SelectNodes("./div[@class='p_item_container p_item_ab ']"); 
                                    string productCode = "";
                                    string productName = "";
                                    string detailPageUrl = "";
                                    string detailPageName = "";

                                    HtmlNode nameNode = itemNode.SelectSingleNode("./div[2]/div[@class=\"listDescript\"]/a[1]");
                                    detailPageUrl = nameNode.Attributes["href"].Value;
                                    int startIndex = detailPageUrl.LastIndexOf("/") + 1;

                                    //商品类型为礼品卡时，length==0，不用获取详情页 
                                    detailPageName = detailPageUrl.Substring(startIndex);
                                    productCode = detailPageName;
                                    productName = nameNode.InnerText.Trim();
                                    detailPageName = category3Code + "_" + detailPageName;

                                    if (!goodsDic.ContainsKey(detailPageName))
                                    {
                                        goodsDic.Add(detailPageName, null);
                                        Dictionary<string, string> p2vs = new Dictionary<string, string>();
                                        p2vs.Add("detailPageUrl", detailPageUrl);
                                        p2vs.Add("detailPageName", detailPageName);
                                        p2vs.Add("productCode", productCode);
                                        p2vs.Add("productName", productName);
                                        p2vs.Add("category1Code", category1Code);
                                        p2vs.Add("category2Code", category2Code);
                                        p2vs.Add("category3Code", category3Code);
                                        p2vs.Add("category1Name", category1Name);
                                        p2vs.Add("category2Name", category2Name);
                                        p2vs.Add("category3Name", category3Name);
                                        resultEW.AddRow(p2vs);
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
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