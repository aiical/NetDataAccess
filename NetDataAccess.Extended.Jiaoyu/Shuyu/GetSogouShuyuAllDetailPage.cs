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

namespace NetDataAccess.Extended.Jiaoyu.Shuyu
{
    public class GetSogouShuyuAllDetailPage : ExternalRunWebPage
    { 
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            try
            {
                //this.GetList(listSheet);
                this.GetWordCount(listSheet);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return true;
        }

        private void GetList(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("cate1", 0);
            resultColumnDic.Add("cateId1", 1);
            resultColumnDic.Add("cate2", 2);
            resultColumnDic.Add("cateId2", 3);
            resultColumnDic.Add("cate3", 4);
            resultColumnDic.Add("cateId3", 5);
            resultColumnDic.Add("name", 6);
            resultColumnDic.Add("wordCount", 7);
            string resultFilePath = Path.Combine(exportDir, "教育_术语_scel详情信息.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"];

                string cate1 = row["cate1"];
                string cate2 = row["cate2"];
                string cateId1 = row["cateId1"];
                string cateId2 = row["cateId2"];
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string localFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir);

                    try
                    {
                        HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                        HtmlNodeCollection itemNodes = htmlDoc.DocumentNode.SelectNodes("//div[@class=\"dict_info_list\"]/ul/li");

                        HtmlNode itemNode = itemNodes[0];
                        string text = itemNode.InnerText.Trim();
                        int splitBeginIndex = text.IndexOf("：");
                        int splitEndIndex = text.IndexOf("个");
                        int wordCount = int.Parse(text.Substring(splitBeginIndex + 1, splitEndIndex - splitBeginIndex - 1));

                        Dictionary<string, string> f2vs = new Dictionary<string, string>();
                        f2vs.Add("cate1", row["cate1"]);
                        f2vs.Add("cateId1", row["cateId1"]);
                        f2vs.Add("cate2", row["cate2"]);
                        f2vs.Add("cateId2", row["cateId2"]);
                        f2vs.Add("cate3", row["cate3"]);
                        f2vs.Add("cateId3", row["cateId3"]);
                        f2vs.Add("name", row["name"]);
                        f2vs.Add("wordCount", wordCount.ToString());
                        resultEW.AddRow(f2vs);
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
            }
            resultEW.SaveToDisk();
        }
        private void GetWordCount(IListSheet listSheet)
        {
            string[] inputParameters = this.Parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            String textFileDir = inputParameters[0];
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("cate1", 0);
            resultColumnDic.Add("cateId1", 1);
            resultColumnDic.Add("cate2", 2);
            resultColumnDic.Add("cateId2", 3);
            resultColumnDic.Add("cate3", 4);
            resultColumnDic.Add("cateId3", 5);
            resultColumnDic.Add("name", 6);
            resultColumnDic.Add("fileName", 7);
            resultColumnDic.Add("wordCountInPage", 8);
            resultColumnDic.Add("wordCountInFile", 9);
            string resultFilePath = Path.Combine(exportDir, "教育_术语_scel词数.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"];

                string cate1 = row["cate1"];
                string cate2 = row["cate2"];
                string cateId1 = row["cateId1"];
                string cateId2 = row["cateId2"];
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string localFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir);

                    try
                    {
                        HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                        HtmlNodeCollection itemNodes = htmlDoc.DocumentNode.SelectNodes("//div[@class=\"dict_info_list\"]/ul/li");

                        HtmlNode itemNode = itemNodes[0];
                        string text = itemNode.InnerText.Trim();
                        int splitBeginIndex = text.IndexOf("：");
                        int splitEndIndex = text.IndexOf("个");
                        int wordCountInPage = int.Parse(text.Substring(splitBeginIndex + 1, splitEndIndex - splitBeginIndex - 1));


                        int fileIdBeginIndex = detailUrl.LastIndexOf("/");
                        string fileId = detailUrl.Substring(fileIdBeginIndex + 1);

                        string fileName = CommonUtil.ProcessFileName(row["name"] + "_" + fileId + ".scel.txt", "_");

                        string textFilePath = Path.Combine(textFileDir, fileName);
                        string textInfo = FileHelper.GetTextFromFile(textFilePath);
                        string[] words = textInfo.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
                        int wordCountInFile = words.Length;


                        Dictionary<string, string> f2vs = new Dictionary<string, string>();
                        f2vs.Add("cate1", row["cate1"]);
                        f2vs.Add("cateId1", row["cateId1"]);
                        f2vs.Add("cate2", row["cate2"]);
                        f2vs.Add("cateId2", row["cateId2"]);
                        f2vs.Add("cate3", row["cate3"]);
                        f2vs.Add("cateId3", row["cateId3"]);
                        f2vs.Add("name", row["name"]);
                        f2vs.Add("fileName", fileName);
                        f2vs.Add("wordCountInPage", wordCountInPage.ToString());
                        f2vs.Add("wordCountInFile", wordCountInFile.ToString());
                        resultEW.AddRow(f2vs);
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
            }
            resultEW.SaveToDisk();
        } 
        private string getNameFromFullName(string fullCateName)
        {
            int splitBeginIndex = fullCateName.IndexOf("(");
            return splitBeginIndex < 0 ? fullCateName : fullCateName.Substring(0, splitBeginIndex);
        }

        private int getPageCountFromFullName(string fullName)
        {
            int splitBeginIndex = fullName.IndexOf("(");
            int splitEndIndex = fullName.IndexOf(")");
            return int.Parse(fullName.Substring(splitBeginIndex + 1, splitEndIndex - splitBeginIndex - 1)) / 10 + 1;
        }

    }
}