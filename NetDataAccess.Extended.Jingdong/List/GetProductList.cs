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
using System.Drawing;
using NetDataAccess.Base.Reader; 
using NetDataAccess.AppAccessBase;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Android;
using System.Collections.ObjectModel;
using System.Xml;

namespace NetDataAccess.Extended.Jingdong.List
{
    /// <summary>
    /// GetProductList
    /// </summary>
    [System.Runtime.InteropServices.ComVisibleAttribute(true)]
    public class GetProductList : ExternalRunWebPage
    { 
        #region 初始化与手机App的连接
        /// <summary>
        /// 初始化与手机App的连接
        /// </summary>
        private AndroidAppAccess InitAppAccess()
        {
            AndroidAppAccess appAccess = new AndroidAppAccess();
            Dictionary<string, string> initParams = new Dictionary<string, string>();
            initParams.Add("deviceName", "192.168.124.101:5555");
            initParams.Add("platformVersion", "4.3");
            initParams.Add("appPackage", "com.jingdong.app.mall");
            initParams.Add("appActivity", "com.jingdong.app.mall.main.MainActivity"); 
            initParams.Add("url", "http://127.0.0.1:4723/wd/hub");
            appAccess.InitDriver(initParams);

            //留出五秒钟的实际，等待手机被唤醒，及app被启动
            Thread.Sleep(5000);

            return appAccess;
        }

        public override bool BeforeAllGrab()
        { 
            return base.BeforeAllGrab();
        } 
        #endregion

        #region 关闭手机连接 
        private void CloseAppAccess(AndroidAppAccess appAccess)
        {
            if (appAccess != null)
            {
                appAccess.Close();
            }
        }
        #endregion

        public override void GetDataByOtherAccessType(Dictionary<string, string> listRow)
        {
            string shopName = listRow["name"];
            string keywordShopFilePath = this.GetShopProductFilePath(shopName);
            if (!File.Exists(keywordShopFilePath))
            {
                this.RunPage.InvokeAppendLogText("开始连接手机APP", LogLevelType.System, true);
                AndroidAppAccess appAccess = this.InitAppAccess();
                this.RunPage.InvokeAppendLogText("连接手机APP成功", LogLevelType.System, true);

                try
                {
                    this.GotoShopListPage(appAccess, shopName);
                    List<Dictionary<string, string>> allProducts = new List<Dictionary<string, string>>();
                    Dictionary<string, string> allProductNames = new Dictionary<string, string>();
                    this.GetShopProductInfos(appAccess, shopName, allProducts, allProductNames, 0);
                    this.SaveShopProductInfoToLocalFile(keywordShopFilePath, allProducts, shopName);
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                finally
                {
                    this.CloseAppAccess(appAccess);
                }
            }
        }
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            string filePath = Path.Combine(this.RunPage.GetExportDir(), "京东App店铺商品列表.xlsx");
            Dictionary<string, int> columnNameToIndex = CommonUtil.InitStringIndexDic(new string[]{
                "shopName",
                "productName", 
                "price",  
                "commentNum",   
                "goodMark"});

            ExcelWriter ew = new ExcelWriter(filePath, "List", columnNameToIndex);

            int keywordCount = listSheet.RowCount;
            Dictionary<string, string> productNameDic = new Dictionary<string, string>();
            for (int i = 0; i < keywordCount; i++)
            {
                Dictionary<string, string> listRow = listSheet.GetRow(i);
                string shopName = listRow["shopName"];
                string shopFilePath = this.GetShopProductFilePath(shopName);
                List<Dictionary<string, string>> shopProductInfos = this.ReadKeywordShopFromLocalFile(shopFilePath);
                foreach (Dictionary<string, string> shopProductInfo in shopProductInfos)
                {
                    string productName = shopProductInfo["productName"];
                    if (!productNameDic.ContainsKey(productName))
                    {
                        productNameDic.Add(productName, null);
                        shopProductInfo.Add("shopName", shopName); 
                        ew.AddRow(shopProductInfo);
                    }
                }
            }
            ew.SaveToDisk(); 

            return base.AfterAllGrab(listSheet);
        }

        private void GetShopProductInfos(AndroidAppAccess appAccess, string keyword, List<Dictionary<String, string>> allShopProducts, Dictionary<string, string> allProductNames, int tryGotTime)
        {
            try
            {
                ReadOnlyCollection<AndroidElement> allProductElements = appAccess.GetElementsById("com.jd.lib.jshop:id/product_list_item", true);
                XmlElement rootElement = appAccess.GetXmlRootElement();
                XmlNodeList productXmlElements = rootElement.SelectNodes("//android.widget.RelativeLayout[@resource-id=\"com.jd.lib.jshop:id/product_list_item\"]");
                bool gotNewProduct = false;
                for (int i = 0; i < allProductElements.Count; i++)
                {
                    AndroidElement productElement = allProductElements[i];
                    XmlElement productXmlElement = (XmlElement)productXmlElements[i];
                    if (this.GetProductInfo(appAccess, productElement, productXmlElement, allShopProducts, allProductNames))
                    {
                        gotNewProduct = true;
                    }
                    /*
                    if (i + 1 < shopXmlElements.Count)
                    {
                        allShopElements = appAccess.GetElementsById("com.jingdong.app.mall:id/b6j", true);
                        rootElement = appAccess.GetXmlRootElement();
                        shopXmlElements = rootElement.SelectNodes("//android.widget.LinearLayout[@resource-id=\"com.jingdong.app.mall:id/b6j\"]");
                    }
                     */
                }
                if (gotNewProduct)
                {
                    tryGotTime = 0;
                }
                else
                {
                    tryGotTime++;
                }

                //如果没有到底部，那么继续滑动
                if (tryGotTime < 3)
                {
                    appAccess.Swipe(new Point(100, 1800), new Point(100, 100), 1000);
                    this.GetShopProductInfos(appAccess, keyword, allShopProducts, allProductNames, tryGotTime);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private bool GetProductInfo(AndroidAppAccess appAccess, AndroidElement productElement, XmlElement productXmlElement, List<Dictionary<String, string>> allShopProducts, Dictionary<string, string> allProductNames)
        {
            try
            {
                XmlElement nameXmlElement = (XmlElement)productXmlElement.SelectSingleNode("./descendant::android.widget.TextView[@resource-id=\"com.jd.lib.jshop:id/product_item_name\"]");
                if (nameXmlElement != null)
                {
                    string productName = nameXmlElement.GetAttribute("text");
                    productName = productName.StartsWith("1 ") ? productName.Substring(2).Trim() : productName.Trim();
                    if (!allProductNames.ContainsKey(productName))
                    {
                        XmlElement priceXmlElement = (XmlElement)productXmlElement.SelectSingleNode("./descendant::android.widget.TextView[@resource-id=\"com.jd.lib.jshop:id/product_item_jdPrice\"]");
                        if (priceXmlElement != null)
                        {
                            string price = priceXmlElement.GetAttribute("text");

                            XmlElement commentNumXmlElement = (XmlElement)productXmlElement.SelectSingleNode("./descendant::android.widget.TextView[@resource-id=\"com.jd.lib.jshop:id/product_item_commentNumber\"]");
                            string commentNum = commentNumXmlElement == null ? "" : commentNumXmlElement.GetAttribute("text");

                            XmlElement goodMarkXmlElement = (XmlElement)productXmlElement.SelectSingleNode("./descendant::android.widget.TextView[@resource-id=\"com.jd.lib.jshop:id/product_item_good\"]");
                            string goodMark = goodMarkXmlElement == null ? "" : goodMarkXmlElement.GetAttribute("text");

                            allProductNames.Add(productName, null);
                            Dictionary<string, string> shopInfo = new Dictionary<string, string>();
                            shopInfo.Add("productName", productName);
                            shopInfo.Add("commentNum", commentNum);
                            shopInfo.Add("price", price);
                            shopInfo.Add("goodMark", goodMark);
                            allShopProducts.Add(shopInfo);
                            return true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
            }
            return false;
        }

        private List<Dictionary<string, string>> ReadKeywordShopFromLocalFile(string filePath)
        {
            List<Dictionary<string, string>> rows = new List<Dictionary<string, string>>();
            CsvReader cr = new CsvReader(filePath);
            int rowCount = cr.GetRowCount();
            for (int i = 0; i < rowCount; i++)
            {
                Dictionary<string, string> row = cr.GetFieldValues(i);
                rows.Add(row);
            }
            return rows;
        }

        private void SaveShopProductInfoToLocalFile(string filePath, List<Dictionary<string, string>> allProducts, string shopName)
        {
            Dictionary<string, int> columnNameToIndex = new Dictionary<string, int>();
            columnNameToIndex.Add("shopName", 0);
            columnNameToIndex.Add("productName", 1);
            columnNameToIndex.Add("price", 2);
            columnNameToIndex.Add("commentNum", 3);
            columnNameToIndex.Add("goodMark", 4);  
            CsvWriter cw = new CsvWriter(filePath, columnNameToIndex);
            foreach (Dictionary<string, string> skuInfo in allProducts)
            {
                skuInfo.Add("shopName", shopName);
                cw.AddRow(skuInfo);
            }
            cw.SaveToDisk();
        }

        private string GetShopProductFilePath(string shopName)
        {
            string filePath = Path.Combine(this.RunPage.GetExportDir(), CommonUtil.ProcessFileName(shopName, "_") + ".csv");
            return filePath;
        }

        private void GotoShopListPage(AndroidAppAccess appAccess, string shopName)
        {
            AndroidElement unknownElement = appAccess.GetElementByIds(new string[] { "com.jingdong.app.mall:id/aqz", "com.jingdong.app.mall:id/a10" }, true);
            if (unknownElement.TagName == "android.widget.EditText")
            {
                unknownElement.Click();
            }
            else
            {
                //如果获取到的是是关闭广告按钮
                unknownElement.Click();

                //点击输入框
                AndroidElement gotoInputElement = appAccess.GetElementById("com.jingdong.app.mall:id/a10", true);
                gotoInputElement.Click();
            }

            /*
            //点击后进入查询页面
            AndroidElement gotoInputElement = null;

            try
            {
                gotoInputElement = appAccess.GetElementById("com.jingdong.app.mall:id/a10", true);
            }
            catch (Exception ex)
            {
                AndroidElement adElement = appAccess.GetElementById("com.jingdong.app.mall:id/aqz", true);
                adElement.Click();
                gotoInputElement = appAccess.GetElementById("com.jingdong.app.mall:id/a10", true);
            }

            gotoInputElement.Click();
             */

            //录入店铺keyword
            AndroidElement inputKeywordElement = appAccess.GetElementById("com.jd.lib.search:id/search_text", true);
            inputKeywordElement.SendKeys(shopName);

            //查询按钮
            AndroidElement searchBtnElement = appAccess.GetElementById("com.jd.lib.search:id/search_btn", true);
            searchBtnElement.Click();

            AndroidElement shopLinkElement = appAccess.GetElementByIds(new string[] { "com.jd.lib.search:id/search_recommend_shop_info", "com.jd.lib.search:id/product_list_shop_first_line", "com.jd.lib.search:id/jshop_list_item_name" }, true);
            shopLinkElement.Click();

            AndroidElement allProductLinkElement = appAccess.GetElementByClassNameAndText("android.widget.TextView", new string[] { "全部商品" }, false, true);
            allProductLinkElement.Click();

            AndroidElement productListElement = appAccess.GetElementById("com.jd.lib.jshop:id/product_list", true);

        }
    }
}