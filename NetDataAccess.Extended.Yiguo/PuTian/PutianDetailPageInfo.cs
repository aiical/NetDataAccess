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
    public class PutianDetailPageInfo : CustomProgramBase
    {
        public bool Run(string parameters, IListSheet listSheet)
        {
            return this.GeneratePageInfo(listSheet);
        }

        private bool GeneratePageInfo(IListSheet listSheet)
        {
            bool succeed = true;
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            string exportDir = this.RunPage.GetExportDir();
            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                "商品名称",
                "商品编码",
                "价格",  
                "规格",
                "产地", 
                "一级分类编码",
                "二级分类编码", 
                "三级分类编码", 
                "一级分类", 
                "二级分类",
                "三级分类",
                "地区",
                "商品页URL"});
            string resultFilePath = Path.Combine(exportDir, "莆田网商品详情" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx");

            Dictionary<string, string> resultColumnFormat = new Dictionary<string, string>();
            resultColumnFormat.Add("价格", "#,##0.00");

            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, resultColumnFormat);

            GetDetailInfo(listSheet, pageSourceDir, resultEW);

            resultEW.SaveToDisk();

            return succeed;
        }

        private void GetDetailInfo(IListSheet listSheet, string pageSourceDir, ExcelWriter resultEW)
        {
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string pageUrl = listSheet.PageUrlList[i];
                string pageName = listSheet.PageNameList[i];
                string localFilePath = this.RunPage.GetFilePath(pageUrl, pageSourceDir);
                string productSysNo = row["productSysNo"];
                string category1Code = row["category1Code"];
                string category2Code = row["category2Code"];
                string category3Code = row["category3Code"];
                string category1Name = row["category1Name"];
                string category2Name = row["category2Name"];
                string category3Name = row["category3Name"];
                string name = "";
                string district = row["district"];
                string cookie = row["cookie"];   
                Nullable<decimal> jiage = null;
                string chandi = "";
                string guige = ""; 

                TextReader tr = null;

                try
                {
                    tr = new StreamReader(localFilePath);
                    string webPageHtml = tr.ReadToEnd().Trim(); 
                    HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                    htmlDoc.LoadHtml(webPageHtml);

                    HtmlNode nameNode = htmlDoc.DocumentNode.SelectSingleNode("//div[@class=\"product-intro product-container\"]/h1");
                    name = CommonUtil.HtmlDecode(nameNode.InnerText.Trim());

                    HtmlNode priceNode = htmlDoc.DocumentNode.SelectSingleNode("//div[@class=\"product-intro product-container\"]/div[@class=\"price\"]/span");
                    if (priceNode != null)
                    {
                        string priceStr = CommonUtil.HtmlDecode(priceNode.InnerText.Trim());
                        decimal testJiage = 0;
                        if (decimal.TryParse(priceStr.Substring(1), out testJiage))
                        {
                            jiage = testJiage;
                        }
                    }

                    //规格、产地、存储方式，无品牌字段，需要从名称里提取
                    HtmlNodeCollection allPropertyNodes = htmlDoc.DocumentNode.SelectNodes("//div[@class=\"product-intro product-container\"]/ul[@class=\"unit-origin-storage\"]/li");
                    if (allPropertyNodes != null)
                    {
                        foreach (HtmlNode propertyNode in allPropertyNodes)
                        {
                            HtmlNode pNameNode = propertyNode.FirstChild;

                            string propertyText = pNameNode.InnerText;
                            if (pNameNode.NextSibling != null)
                            {
                                string propertyValue = CommonUtil.HtmlDecode(pNameNode.NextSibling.InnerText.Trim());
                                if (propertyText.Contains("规格"))
                                {
                                    guige = propertyValue;
                                }
                                else if (propertyText.Contains("产地"))
                                {
                                    chandi = propertyValue;
                                }
                            }
                        }
                    }

                    Dictionary<string, object> f2vs = new Dictionary<string, object>(); 
                    f2vs.Add("一级分类编码", category1Code);
                    f2vs.Add("二级分类编码", category2Code);
                    f2vs.Add("三级分类编码", category3Code);
                    f2vs.Add("一级分类", category1Name);
                    f2vs.Add("二级分类", category2Name);
                    f2vs.Add("三级分类", category3Name); 
                    f2vs.Add("商品编码", productSysNo);
                    f2vs.Add("商品名称", name);
                    f2vs.Add("价格", jiage);  
                    f2vs.Add("产地", chandi);
                    f2vs.Add("规格", guige);
                    f2vs.Add("地区", district);
                    f2vs.Add("商品页URL", pageUrl);
                    resultEW.AddRow(f2vs);
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
    }
}