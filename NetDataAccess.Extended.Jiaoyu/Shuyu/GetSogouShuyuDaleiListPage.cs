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
    public class GetSogouShuyuDaleiListPage : ExternalRunWebPage
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
            resultColumnDic.Add("cate1", 5);
            resultColumnDic.Add("cateId1", 6);
            resultColumnDic.Add("cate2", 7);
            resultColumnDic.Add("cateId2", 8);
            string resultFilePath = Path.Combine(exportDir, "教育_术语_小类.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null); 
            
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"];
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string localFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir);

                    try
                    {
                        HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                        HtmlNodeCollection cate1LinkNodes = htmlDoc.DocumentNode.SelectNodes("//div[@id=\"dict_nav_list\"]/ul/li[contains(@class, \"nav_list\")]/a");

                        for (int j = 0; j < cate1LinkNodes.Count; j++)
                        {
                            HtmlNode cate1LinkNode = cate1LinkNodes[j];
                            string linkUrl = cate1LinkNode.GetAttributeValue("href", "");
                            int linkIdBeginIndex = linkUrl.LastIndexOf("/") + 1;
                            string id = linkUrl.Substring(linkIdBeginIndex).Trim();

                            if (id == "167")
                            {
                                HtmlNodeCollection cateDiquLinkNodes = htmlDoc.DocumentNode.SelectNodes("//div[contains(@class, \"citylistcate\")]/a");
                                foreach (HtmlNode cateDiquLinkNode in cateDiquLinkNodes)
                                {
                                    string linkDiquUrl = cateDiquLinkNode.GetAttributeValue("href", "");
                                    string diqu = cateDiquLinkNode.InnerText.Trim();
                                    int linkDiquIdBeginIndex = linkDiquUrl.LastIndexOf("/") + 1;
                                    string diquId = linkDiquUrl.Substring(linkDiquIdBeginIndex).Trim();

                                    Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                    f2vs.Add("detailPageUrl", "https://pinyin.sogou.com" + linkDiquUrl);
                                    f2vs.Add("detailPageName", diquId);
                                    f2vs.Add("cate1", id);
                                    f2vs.Add("cateId1", id);
                                    f2vs.Add("cate2", diqu);
                                    f2vs.Add("cateId2", diquId);
                                    resultEW.AddRow(f2vs);
                                }
                            }
                            else
                            {

                                Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                f2vs.Add("detailPageUrl", "https://pinyin.sogou.com" + linkUrl);
                                f2vs.Add("detailPageName", id);
                                f2vs.Add("cate1", id);
                                f2vs.Add("cateId1", id);
                                resultEW.AddRow(f2vs);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    } 
                }
            }
            resultEW.SaveToDisk();
        }
    }
}