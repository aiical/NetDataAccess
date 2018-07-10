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
    public class SMDetailPageInfo : CustomProgramBase
    {
        public bool Run(string parameters, IListSheet listSheet)
        {
            return this.GenerateDetailPageInfo(listSheet);
        }

        private bool GenerateDetailPageInfo(IListSheet listSheet)
        {
            bool succeed = true;
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                "商品名称",
                "价格",
                "品牌",
                "一级分类编码", 
                "一级分类", 
                "二级分类编码", 
                "二级分类",
                "三级分类编码", 
                "三级分类",
                "SKU单位", 
                "重量", 
                "url", 
                "商品编码"});


            Dictionary<string, string> resultColumnFormat = new Dictionary<string, string>();
            resultColumnFormat.Add("价格", "#,##0.00");

            string readDetailDir = this.RunPage.GetReadFileDir();
            string exportDir = this.RunPage.GetExportDir();
            string resultFilePath = Path.Combine(exportDir, this.RunPage.Project.Name + "_All.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, resultColumnFormat);

            GetList(listSheet, pageSourceDir, resultEW); 

            resultEW.SaveToDisk(); 

            return succeed;
        }

        private void GetList(IListSheet listSheet, string pageSourceDir, ExcelWriter resultEW)
        { 
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string pageUrl = listSheet.PageUrlList[i];
                string pageName = listSheet.PageNameList[i];
                string localFilePath = this.RunPage.GetFilePath(pageUrl, pageSourceDir);
                string productCode = row["productCode"];
                string productName = CommonUtil.ReplaceAsciiByString(row["productName"]);
                string category1Code = row["category1Code"];
                string category2Code = row["category2Code"];
                string category3Code = row["category3Code"];
                string category1Name = row["category1Name"];
                string category2Name = row["category2Name"];
                string category3Name = row["category3Name"];
                string guiGe = "";
                string zhongliang = "";
                string brand = "";
                decimal productCurrentPrice = 0; 

                TextReader tr = null;

                try
                {
                    tr = new StreamReader(localFilePath, Encoding.GetEncoding("GBK"));
                    string webPageHtml = tr.ReadToEnd();
                    HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                    htmlDoc.LoadHtml(webPageHtml);

                    HtmlNodeCollection propertyLabelNodes = htmlDoc.DocumentNode.SelectNodes("//li[@class=\"proAttrSection cls\"]/label");
                    foreach (HtmlNode plNode in propertyLabelNodes)
                    {
                        if (plNode.InnerText.Contains("重量"))
                        {
                            HtmlNode pvNode = plNode.NextSibling;
                            while (pvNode != null && pvNode.Name != "p")
                            {
                                pvNode = pvNode.NextSibling;
                            }
                            if (pvNode != null)
                            {
                                zhongliang = pvNode.SelectSingleNode("./span").ChildNodes[0].InnerText.Trim();
                            }
                            break;
                        }
                    }

                    HtmlNode priceNode = htmlDoc.DocumentNode.SelectSingleNode("//span[@class=\"curPrice\"]");
                    if (priceNode != null)
                    {
                        string priceStr = priceNode.InnerText.Trim();
                        priceStr = priceStr.Substring(1).Trim();
                        productCurrentPrice = decimal.Parse(priceStr);
                    }

                    HtmlNode guiGeNode = htmlDoc.DocumentNode.SelectSingleNode("//span[@class=\"pt3\"]");
                    if (guiGeNode != null)
                    {
                        guiGe = guiGeNode.InnerText.Replace("/", "").Trim();
                    }

                    HtmlNode brandNode = htmlDoc.DocumentNode.SelectSingleNode("//a[@class=\"minorFunc\"]");
                    if (brandNode != null)
                    {
                        brand = brandNode.Attributes["title"].Value;
                        brand = CommonUtil.ReplaceAsciiByString(brand);
                    } 
                }
                catch (Exception ex)
                {
                    this.RunPage.InvokeAppendLogText("读取出错. url = " + pageUrl + ". " + ex.Message, LogLevelType.Error, true);
                    throw ex;
                }

                Dictionary<string, object> f2vs = new Dictionary<string, object>();
                f2vs.Add("商品编码", productCode);
                f2vs.Add("商品名称", productName);
                f2vs.Add("价格", productCurrentPrice);
                f2vs.Add("SKU单位", guiGe);
                f2vs.Add("重量", zhongliang); 
                f2vs.Add("品牌", brand); 
                f2vs.Add("一级分类", category1Name);
                f2vs.Add("二级分类", category2Name);
                f2vs.Add("三级分类", category3Name);
                f2vs.Add("url", pageUrl);  
                f2vs.Add("一级分类编码", category1Code);
                f2vs.Add("二级分类编码", category2Code);
                f2vs.Add("三级分类编码", category3Code);  

                resultEW.AddRow(f2vs);
            }
        }
    }
}