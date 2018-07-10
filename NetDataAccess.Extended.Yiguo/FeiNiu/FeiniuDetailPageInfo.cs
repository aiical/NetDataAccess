﻿using System;
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
    public class FeiniuDetailPageInfo : CustomProgramBase
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
                "一级分类编码", 
                "一级分类", 
                "二级分类编码", 
                "二级分类",
                "三级分类编码", 
                "三级分类",
                "品牌", 
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
                string pageUrl = row["detailPageUrl"];
                string pageName = row["detailPageName"];
                bool giveUpGrab = row["giveUpGrab"] == "Y";
                string localFilePath = this.RunPage.GetFilePath(pageUrl, pageSourceDir);
                string productCode = row["productCode"];
                string productName = row["productName"];
                string category1Code = row["category1Code"];
                string category2Code = row["category2Code"];
                string category3Code = row["category3Code"];
                string category1Name = row["category1Name"];
                string category2Name = row["category2Name"];
                string category3Name = row["category3Name"];
                string pinpai = "";
                decimal productCurrentPrice = 0;

                if (!giveUpGrab)
                {
                    TextReader tr = null;

                    try
                    {
                        tr = new StreamReader(localFilePath);
                        string webPageHtml = tr.ReadToEnd();
                        HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                        htmlDoc.LoadHtml(webPageHtml); 
                        {
                            HtmlNode propertyParentNode = htmlDoc.DocumentNode.SelectSingleNode("//dl[@class=\"dl-proInfo\"]");
                            if (propertyParentNode != null)
                            { 
                                HtmlNodeCollection propertyNodes = propertyParentNode.SelectNodes("./dd");
                                foreach (HtmlNode pNode in propertyNodes)
                                {
                                    string pText = pNode.InnerText.Trim();
                                    if (pText.StartsWith("品牌："))
                                    {
                                        pinpai = pText.Substring(3);
                                        break;
                                    }
                                }
                            }
                        }
                        {
                            HtmlNode propertyParentNode = htmlDoc.DocumentNode.SelectSingleNode("//ul[@class=\"depict-list fn-clear\"]");
                            if (propertyParentNode != null)
                            { 
                                HtmlNodeCollection propertyNodes = propertyParentNode.SelectNodes("./li");
                                foreach (HtmlNode pNode in propertyNodes)
                                {
                                    HtmlNode pnNode = pNode.SelectSingleNode("./span[1]");
                                    string pnText = pnNode.InnerText.Trim();
                                    if (pnText.StartsWith("品牌："))
                                    {
                                        HtmlNode pvNode = pNode.SelectSingleNode("./span[2]");
                                        pinpai = pvNode == null ? "" : pvNode.InnerText.Trim();
                                        break;
                                    }
                                }
                            }
                        }


                        HtmlNode priceNode = htmlDoc.DocumentNode.SelectSingleNode("//font[@class=\"info-price\"]/b[@class=\"JS-control-price\"]");
                        priceNode = priceNode == null ? htmlDoc.DocumentNode.SelectSingleNode("//div[@id=\"J_product_value\"]/div/strong[@class=\"fn-rmb-num\"]") : priceNode;
                        if (priceNode != null)
                        {
                            string priceStr = priceNode.InnerText.Trim();
                            productCurrentPrice = decimal.Parse(priceStr);
                        }
                        else
                        {
                            //throw new Exception("None price! url = " + pageUrl);
                            this.RunPage.InvokeAppendLogText("None price! url = " + pageUrl, LogLevelType.Error, true);

                            if (tr != null)
                            {
                                tr.Close();
                                tr.Dispose();
                            }
                            File.Delete(localFilePath);
                        }
                    }
                    catch (Exception ex)
                    {
                        this.RunPage.InvokeAppendLogText("读取出错. url = " + pageUrl + ". " + ex.Message, LogLevelType.Error, true);
                        throw ex;
                    }
                    finally
                    {
                        if (tr != null)
                        {
                            tr.Close();
                            tr.Dispose();
                        }
                    }

                    Dictionary<string, object> f2vs = new Dictionary<string, object>();
                    f2vs.Add("商品编码", productCode);
                    f2vs.Add("商品名称", productName);
                    f2vs.Add("价格", productCurrentPrice);
                    f2vs.Add("品牌", pinpai);
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
}