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
    public class GetBaiduBaikeListUrls : ExternalRunWebPage
    { 
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            try
            {
                this.GetList(listSheet); 
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
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("word", 6);
            string resultFilePath = Path.Combine(exportDir, "教育_百度百科_词条页面.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);

            string[] inputParameters = this.Parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            String textFileDir = inputParameters[0];

            string[] textFiles = Directory.GetFiles(textFileDir);
            Dictionary<string, string> wordDic = new Dictionary<string, string>();
            foreach (string textFile in textFiles)
            {
                string text = FileHelper.GetTextFromFile(textFile);
                string[] words = text.Split(new String[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string word in words)
                {
                    if (!wordDic.ContainsKey(word))
                    {
                        wordDic.Add(word, null);
                        Dictionary<string, string> f2vs = new Dictionary<string, string>();
                        f2vs.Add("detailPageUrl", "https://baike.baidu.com/item/" + word);
                        f2vs.Add("detailPageName", word);
                        f2vs.Add("word", word); 
                        resultEW.AddRow(f2vs);
                    }
                }
            }
            resultEW.SaveToDisk();
        }
    }
}