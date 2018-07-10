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
using NetDataAccess.Base.Writer;
using NetDataAccess.Base.Reader;
using NetDataAccess.Base.CsvHelper;
using NetDataAccess.Base.DB;
using HtmlAgilityPack;
using NetDataAccess.Base.Server;

namespace NetDataAccess.Extended.Yiguo
{
    /// <summary>
    /// 本来生活
    /// 从本地html中获取并记录下商品详情信息
    /// </summary>
    public class BenlaiDetailPageInfo : CustomProgramBase
    {
        #region 入口函数
        public bool Run(string parameters, IListSheet listSheet)
        {
            return this.GenerateDetailPageInfo(listSheet);
        }
        #endregion

        #region 生成并输出商品详情
        private bool GenerateDetailPageInfo(IListSheet listSheet)
        {
            bool succeed = true;
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                "detailPageUrl",
                "detailPageName",
                "cookie",
                "grabStatus",
                "giveUpGrab",
                "商品名称",
                "价格",
                "一级分类编码", 
                "一级分类", 
                "二级分类编码", 
                "二级分类",
                "三级分类编码", 
                "三级分类",
                "规格", 
                "温馨提示", 
                "满减", 
                "评论数", 
                "好评", 
                "中评", 
                "差评", 
                "好评度", 
                "url", 
                "地区", 
                "productPromotionWord",
                "原价",
                "商品编码"});


            Dictionary<string, string> resultColumnFormat = new Dictionary<string, string>();
            resultColumnFormat.Add("价格", "#,##0.00");
            resultColumnFormat.Add("评论数", "#,##0");
            resultColumnFormat.Add("好评", "#,##0");
            resultColumnFormat.Add("中评", "#,##0");
            resultColumnFormat.Add("差评", "#,##0");
            resultColumnFormat.Add("好评度", "0.00%");
            resultColumnFormat.Add("原价", "#,##0.00");

            string readDetailDir = this.RunPage.GetReadFileDir();
            string exportDir = this.RunPage.GetExportDir();
            string resultFilePath = Path.Combine(exportDir, "本来生活获取所有详情页库存.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, resultColumnFormat);

            GenerateDetailPageInfo(listSheet, pageSourceDir, resultEW);

            resultEW.SaveToDisk();

            return succeed;
        }
        #endregion

        #region 从本地html中生成并输出商品详情
        private void GenerateDetailPageInfo(IListSheet listSheet, string pageSourceDir, ExcelWriter resultEW)
        {
            string detailPageUrlPrefix = "http://www.benlai.com";
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string pageUrl = listSheet.PageUrlList[i];
                string pageName = listSheet.PageNameList[i];
                string localFilePath = this.RunPage.GetFilePath(pageUrl, pageSourceDir);
                string productSysNo = row["productSysNo"];
                string productName = row["productName"];
                string productPromotionWord = row["productPromotionWord"];
                string productCurrentPrice = row["productCurrentPrice"];
                string productOldPrice = row["productOldPrice"];
                string category1Code = row["category1Code"];
                string category2Code = row["category2Code"];
                string category3Code = row["category3Code"];
                string category1Name = row["category1Name"];
                string category2Name = row["category2Name"];
                string category3Name = row["category3Name"];
                string district = row["district"];
                string cookie = row["cookie"];
                string standard = "";
                int totalCommentCount = 0;
                int hCommentCount = 0;
                int mCommentCount = 0;
                int lCommentCount = 0;
                Nullable<decimal> hPer = null;
                string prompt = "";
                string gifts = "";
                string pricePageUrl = "";

                TextReader tr = null;

                try
                { 
                    tr = new StreamReader(localFilePath);
                    string webPageHtml = tr.ReadToEnd();
                    HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                    htmlDoc.LoadHtml(webPageHtml);
                     
                    HtmlNodeCollection propertyNodes = htmlDoc.DocumentNode.SelectNodes("//*[@id=\"_ProductDetails\"]/div[@class=\"good15_intro\"]/div[3]/dl");
                    foreach (HtmlNode pNode in propertyNodes)
                    {
                        HtmlNode dtNode = pNode.SelectSingleNode("./dt");
                        if (dtNode != null)
                        {
                            if (dtNode.InnerText.Contains("规格"))
                            {
                                HtmlNode ddNode = pNode.SelectSingleNode("./dd");
                                if (ddNode != null)
                                {
                                    standard = ddNode.InnerText.Trim();
                                }
                                break;
                            }
                        }
                    }
                     
                    HtmlNodeCollection allCommentNodes = htmlDoc.DocumentNode.SelectNodes("//*[@id=\"reviewPage\"]/p");

                    HtmlNode totalNode = allCommentNodes[0];
                    string[] totalSplits = totalNode.InnerText.Trim().Split(new string[] { "(", "条" }, StringSplitOptions.RemoveEmptyEntries);
                    totalCommentCount = int.Parse(totalSplits[1]);

                    HtmlNode hNode = allCommentNodes[1];
                    string[] hSplits = hNode.InnerText.Trim().Split(new string[] { "(", "条" }, StringSplitOptions.RemoveEmptyEntries);
                    hCommentCount = int.Parse(hSplits[1]);

                    HtmlNode mNode = allCommentNodes[2];
                    string[] mSplits = mNode.InnerText.Trim().Split(new string[] { "(", "条" }, StringSplitOptions.RemoveEmptyEntries);
                    mCommentCount = int.Parse(mSplits[1]);

                    HtmlNode lNode = allCommentNodes[3];
                    string[] lSplits = lNode.InnerText.Trim().Split(new string[] { "(", "条" }, StringSplitOptions.RemoveEmptyEntries);
                    lCommentCount = int.Parse(lSplits[1]);

                    hPer = totalCommentCount == 0 ? null : (Nullable<decimal>)(((decimal)hCommentCount) / (decimal)totalCommentCount);

                    //温馨提示
                    HtmlNode promptNode = htmlDoc.DocumentNode.SelectSingleNode("//*[@class=\"good15_prompt\"]");
                    if (promptNode != null)
                    {
                        prompt = promptNode.InnerText.Trim().Replace(" ", "");
                        if (prompt.StartsWith("温馨提示："))
                        {
                            prompt = prompt.Substring(5);
                        }
                    }

                    //满减
                    HtmlNode giftsNode = htmlDoc.DocumentNode.SelectSingleNode("//*[@class=\"intro_gifts\"]");
                    if (giftsNode != null)
                    {
                        gifts = giftsNode.InnerText.Trim().Replace(" ", "");
                    } 
                }
                catch (Exception ex)
                {
                    this.RunPage.InvokeAppendLogText("读取出错. url = " + pageUrl + ". " + ex.Message, LogLevelType.Error, true);
                    throw ex;
                }

                Dictionary<string, object> f2vs = new Dictionary<string, object>();
                f2vs.Add("商品名称", productName);
                if (!CommonUtil.IsNullOrBlank(productCurrentPrice))
                {
                    f2vs.Add("价格", decimal.Parse(productCurrentPrice));
                }
                f2vs.Add("一级分类", category1Name);
                f2vs.Add("二级分类", category2Name);
                f2vs.Add("三级分类", category3Name);
                f2vs.Add("规格", standard);
                f2vs.Add("评论数", totalCommentCount);
                f2vs.Add("好评", hCommentCount);
                f2vs.Add("中评", mCommentCount);
                f2vs.Add("差评", lCommentCount);
                if (hPer != null)
                {
                    f2vs.Add("好评度", hPer);
                }
                f2vs.Add("url", pageUrl);
                f2vs.Add("地区", district);
                f2vs.Add("productPromotionWord", productPromotionWord);
                if (!CommonUtil.IsNullOrBlank(productOldPrice))
                {
                    f2vs.Add("原价", decimal.Parse(productOldPrice));
                }
                f2vs.Add("商品编码", productSysNo);
                f2vs.Add("一级分类编码", category1Code);
                f2vs.Add("二级分类编码", category2Code);
                f2vs.Add("三级分类编码", category3Code);
                f2vs.Add("满减", gifts);
                f2vs.Add("温馨提示", prompt);
                f2vs.Add("detailPageName", pageName);
                f2vs.Add("cookie", cookie);
                string prefix = "";
                switch (district)
                {
                    case "华东":
                        //prefix = "/huadong";
                        prefix = "";
                        break;
                    case "华北":
                        prefix = "";
                        break;
                    default:
                        break;
                }
                pricePageUrl = detailPageUrlPrefix + prefix + "/ajax/GetProductPrice?SysNo=" + productSysNo + "&_=1451335148388";
                f2vs.Add("detailPageUrl", pricePageUrl);

                resultEW.AddRow(f2vs);
            }
        }
        #endregion
    }
}