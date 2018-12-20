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

namespace NetDataAccess.Extended.LiShi.BaiDuBaiKe
{
    public class GetBaiKeNianFenPages : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        { 
            this.GetYearInfos(listSheet);
            return true;
        }

        public override void CheckRequestCompleteFile(string webPageText, Dictionary<string, string> listRow)
        {
            string yearName = listRow["yearName"]; 
            if (webPageText.Contains("您所访问的页面不存在"))
            {
                throw new GiveUpException("访问的页面不存在, yearName = " + yearName);
            } 
        }

        public override bool AfterGrabOneCatchException(string pageUrl, Dictionary<string, string> listRow, Exception ex)
        {
            this.RunPage.InvokeAppendLogText("准备放弃爬取" + (ex is GiveUpException || ex.InnerException is GiveUpException ? "GiveUpException" : "Exception"), LogLevelType.Error, true);
            return ex is GiveUpException || ex.InnerException is GiveUpException;
        }

        private void GetYearInfos(IListSheet listSheet)
        {
            string sourceDir = this.RunPage.GetDetailSourceFileDir();

            ExcelWriter resultEW = this.CreateResultWriter();
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> listRow = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(listRow[SysConfig.GiveUpGrabFieldName]);
                string detailPageUrl = listRow[SysConfig.DetailPageUrlFieldName];
                if (!giveUp)
                {
                    try
                    {
                        string localFilePath = this.RunPage.GetFilePath(detailPageUrl, sourceDir);
                        string html = FileHelper.GetTextFromFile(localFilePath, Encoding.UTF8);
                        if (!html.Contains("您所访问的页面不存在"))
                        {
                            HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                            htmlDoc.LoadHtml(html);
                            HtmlNode mainInfoNode = htmlDoc.DocumentNode.SelectSingleNode("//div[@class=\"lemma-summary\"]");
                            if (mainInfoNode == null)
                            {
                                this.RunPage.InvokeAppendLogText("此词条不存在摘要信息, pageUrl = " + detailPageUrl, LogLevelType.Error, true);
                            }
                            else
                            { 
                                HtmlNode itemBaseInfoNode = htmlDoc.DocumentNode.SelectSingleNode("//div[@class=\"lemmaWgt-promotion-rightPreciseAd\"]");
                                string itemId = itemBaseInfoNode.GetAttributeValue("data-lemmaid", "");
                                string itemName = itemBaseInfoNode.GetAttributeValue("data-lemmatitle", "");

                                string mainInfo = CommonUtil.HtmlDecode(mainInfoNode.InnerText).Trim();

                                Dictionary<string, string> newRow = new Dictionary<string, string>();
                                newRow.Add("url", detailPageUrl);
                                newRow.Add("yearValue", listRow["yearValue"]);
                                newRow.Add("yearName", listRow["yearName"]);
                                newRow.Add("itemId", itemId);
                                newRow.Add("itemName", itemName);
                                newRow.Add("mainInfo", mainInfo); 
                                resultEW.AddRow(newRow);
                            }

                        }
                        else
                        {
                            this.RunPage.InvokeAppendLogText("放弃解析此页, 所访问的页面不存在, pageUrl = " + detailPageUrl, LogLevelType.Error, true);
                        }
                    }
                    catch (Exception ex)
                    {
                        this.RunPage.InvokeAppendLogText(ex.Message + ". 解析出错， pageUrl = " + detailPageUrl, LogLevelType.Error, true);
                        throw ex;
                    }
                }
            }

            resultEW.SaveToDisk();
        } 
        private ExcelWriter CreateResultWriter()
        {
            String exportDir = this.RunPage.GetExportDir();
            string resultFilePath = Path.Combine(exportDir, "百度百科_年份_摘要信息.xlsx");

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("url", 0);
            resultColumnDic.Add("yearValue", 1);
            resultColumnDic.Add("yearName", 2);
            resultColumnDic.Add("itemId", 3);
            resultColumnDic.Add("itemName", 4);
            resultColumnDic.Add("mainInfo", 5);
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        } 
    }
}