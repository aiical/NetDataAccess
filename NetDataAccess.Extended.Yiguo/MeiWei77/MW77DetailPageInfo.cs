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

namespace NetDataAccess.Extended.Yiguo
{
    /// <summary>
    /// 美味七七
    /// 获取所有商品详情信息
    /// </summary>
    public class MW77DetailPageInfo : CustomProgramBase
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
            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("商品名称", 0);
            resultColumnDic.Add("价格", 1);
            resultColumnDic.Add("一级分类编码", 2);
            resultColumnDic.Add("一级分类", 3);
            resultColumnDic.Add("二级分类编码", 4);
            resultColumnDic.Add("二级分类", 5);
            resultColumnDic.Add("三级分类编码", 6);
            resultColumnDic.Add("三级分类", 7);
            resultColumnDic.Add("规格", 8);
            resultColumnDic.Add("评论数", 9);
            resultColumnDic.Add("满意度", 10);
            resultColumnDic.Add("url", 11); 
            resultColumnDic.Add("原价", 12);
            resultColumnDic.Add("商品编码", 13);


            Dictionary<string, string> resultColumnFormat = new Dictionary<string, string>();
            resultColumnFormat.Add("价格", "#,##0.00");
            resultColumnFormat.Add("评论数", "#,##0");
            resultColumnFormat.Add("满意度", "#,##0.0");
            resultColumnFormat.Add("原价", "#,##0.00");

            string readDetailDir = this.RunPage.GetReadFileDir();
            string exportDir = this.RunPage.GetExportDir();
            string resultFilePath = Path.Combine(exportDir, "美味77商品详情" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, resultColumnFormat);

            GetList(listSheet, pageSourceDir, resultEW); 

            resultEW.SaveToDisk(); 

            return succeed;
        }
        #endregion

        #region 从本地html中生成并输出商品详情
        private void GetList(IListSheet listSheet, string pageSourceDir, ExcelWriter resultEW)
        {
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                if (row["giveUpGrab"] != "是")
                {
                    string pageUrl = listSheet.PageUrlList[i];
                    string localFilePath = this.RunPage.GetFilePath(pageUrl, pageSourceDir);
                    string productCode = row["productCode"];
                    string productName = row["productName"];
                    string productCurrentPrice = row["productCurrentPrice"];
                    string productOldPrice = row["productOldPrice"];
                    string category1Code = row["category1Code"];
                    string category2Code = row["category2Code"];
                    string category3Code = row["category3Code"];
                    string category1Name = row["category1Name"];
                    string category2Name = row["category2Name"];
                    string category3Name = row["category3Name"];
                    string standard = row["standard"];
                    int totalCommentCount = 0;
                    Nullable<decimal> hPer = null;

                    TextReader tr = null;

                    try
                    {
                        tr = new StreamReader(localFilePath);
                        string webPageHtml = tr.ReadToEnd();
                        HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                        htmlDoc.LoadHtml(webPageHtml);

                        HtmlNode commentNode = htmlDoc.DocumentNode.SelectSingleNode("//div[@class=\"comment\"]");
                        HtmlNode pointNode = commentNode.SelectSingleNode("./div[@class=\"point\"]");
                        if (pointNode != null)
                        {
                            string str = pointNode.InnerText.Trim().Replace(" ", "");
                            hPer = decimal.Parse(str.Substring(0, str.Length - 1));
                        }
                        HtmlNode countNode = commentNode.SelectSingleNode("./div[@class=\"count\"]/font");
                        if (countNode != null)
                        {
                            string str = countNode.InnerText.Trim().Replace(" ", "");
                            totalCommentCount = int.Parse(str);
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
                    if (hPer != null)
                    {
                        f2vs.Add("满意度", hPer);
                    }
                    f2vs.Add("url", pageUrl);
                    if (!CommonUtil.IsNullOrBlank(productOldPrice))
                    {
                        f2vs.Add("原价", decimal.Parse(productOldPrice));
                    }
                    f2vs.Add("商品编码", productCode);
                    f2vs.Add("一级分类编码", category1Code);
                    f2vs.Add("二级分类编码", category2Code);
                    f2vs.Add("三级分类编码", category3Code);
                    resultEW.AddRow(f2vs);
                }
            }
        }
        #endregion
    }
}