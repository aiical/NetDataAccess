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

namespace NetDataAccess.Extended.Taobao.List
{
    /// <summary>
    /// GetAllListPage
    /// </summary>
    [System.Runtime.InteropServices.ComVisibleAttribute(true)]
    public class GetShopProductList : ExternalRunWebPage
    {
        #region ExportDir
        private string _ExportDir = null;
        private string ExportDir
        {
            get
            {
                return _ExportDir;
            }
        }
        #endregion 

        #region 初始化与手机App的连接
        /// <summary>
        /// 初始化与手机App的连接
        /// </summary>
        private AndroidAppAccess InitAppAccess()
        {
            AndroidAppAccess appAccess = new AndroidAppAccess();
            Dictionary<string, string> initParams = new Dictionary<string, string>();
            initParams.Add("deviceName", "192.168.124.101:5555");
            initParams.Add("platformVersion", "6.0");
            initParams.Add("appPackage", "com.taobao.taobao");
            initParams.Add("appActivity", "com.taobao.tao.welcome.Welcome"); 
            initParams.Add("url", "http://127.0.0.1:4723/wd/hub");
            appAccess.InitDriver(initParams);

            //留出五秒钟的实际，等待手机被唤醒，及app被启动
            Thread.Sleep(5000);

            return appAccess;
        }

        public override bool BeforeAllGrab()
        {
            try
            {
                string[] allParameters = this.Parameters.Split(new string[] { "," }, StringSplitOptions.None);
                this._ExportDir = allParameters[0];

            }
            catch (Exception ex)
            {
                throw ex;
            } 
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

        public override void GetDataByOtherAcessType(Dictionary<string, string> listRow)
        {
            string shopWebPageUrl = listRow["店铺网址"];
            string shopName = listRow["店铺名"];

            this.RunPage.InvokeAppendLogText("开始连接手机APP", LogLevelType.System, true);
            AndroidAppAccess appAccess = this.InitAppAccess();
            this.RunPage.InvokeAppendLogText("连接手机APP成功", LogLevelType.System, true);

            string shopFilePath = this.GetShopLocalFilePath(shopName, shopWebPageUrl);
            if (!File.Exists(shopFilePath))
            {
                try
                {
                    this.GotoShopSkuPage(appAccess, shopWebPageUrl, shopName);
                    List<Dictionary<String, string>> allSkuInfos = new List<Dictionary<string, string>>();
                    this.GetShopAllSkuInfos(appAccess, shopName, shopWebPageUrl, allSkuInfos);
                    this.SaveShopToLocalFile(shopFilePath, allSkuInfos);
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

        private void GetShopAllSkuInfos( AndroidAppAccess appAccess, string shopName, string shopWebPageUrl, List<Dictionary<String, string>> allSkuInfos)
        {
            ReadOnlyCollection<AndroidElement> allSkuElements = appAccess.GetElementsById("com.taobao.taobao:id/auction_layout", true);
            XmlElement rootElement = appAccess.GetXmlRootElement();
            XmlNodeList skuXmlElements = rootElement.SelectNodes("//android.widget.RelativeLayout[@resource-id=\"com.taobao.taobao:id/auction_layout\"]");

            for (int i = 0; i < skuXmlElements.Count; i++)
            {

                AndroidElement skuElement = allSkuElements[i];
                XmlElement skuXmlElement = (XmlElement)skuXmlElements[i];
                Dictionary<string, string> skuInfo = this.GetSkuInfo(appAccess, shopName, shopWebPageUrl, skuElement, skuXmlElement);
                if (skuInfo != null)
                {
                    allSkuInfos.Add(skuInfo);
                }
                if (i + 1 < skuXmlElements.Count)
                {
                    allSkuElements = appAccess.GetElementsById("com.taobao.taobao:id/auction_layout", true);
                    rootElement = appAccess.GetXmlRootElement();
                    skuXmlElements = rootElement.SelectNodes("//android.widget.RelativeLayout[@resource-id=\"com.taobao.taobao:id/auction_layout\"]");
                }
            }

            //判断是否已经到了最底部
            if (!appAccess.CheckContainTextByXPath("//android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.FrameLayout/android.view.View/android.support.v4.view.ViewPager/android.widget.FrameLayout/android.widget.RelativeLayout/android.support.v7.widget.RecyclerView/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.TextView", new string[] { "没有更多宝贝了" }, false))
            {
                //如果没有到底部，那么继续滑动
                appAccess.Swipe(new Point(100, 1200), new Point(100, 100), 2000);
                this.GetShopAllSkuInfos(appAccess, shopName, shopWebPageUrl, allSkuInfos);
            }
        }

        private Dictionary<string, string> GetSkuInfo(AndroidAppAccess appAccess, string shopName, string shopWebPageUrl, AndroidElement skuElement, XmlElement skuXmlElement)
        {
            string skuName = "";
            string pricePrefix = "";
            AndroidElement nameElement = (AndroidElement)appAccess.GetElementByIdNoWaiting(skuElement, "com.taobao.taobao:id/title", false);
            if (nameElement != null)
            {
                AndroidElement priceElement = (AndroidElement)appAccess.GetElementByIdNoWaiting(skuElement, "com.taobao.taobao:id/priceBlock", false);
                if (priceElement != null)
                {
                    XmlElement priceXmlElement = (XmlElement)skuXmlElement.SelectSingleNode("./android.view.View[@resource-id=\"com.taobao.taobao:id/priceBlock\"]");
                    if (priceXmlElement != null)
                    {
                        skuName = nameElement.Text;
                        string priceStr = priceXmlElement.GetAttribute("content-desc").Trim();
                        int yuanIndex = priceStr.IndexOf("元");
                        pricePrefix = priceStr.Substring(0, yuanIndex).Trim();
                    }
                }
            }

            if (skuName.Length > 0)
            {
                //判断此sku信息是否抓取完成
                string skuFilePath = this.GetSkuInfoLocalFilePath(shopName, skuName);
                Dictionary<string, string> skuInfo = null;
                if (File.Exists(skuFilePath))
                {
                    skuInfo = this.ReadSkuInfoFromLocalFile(skuFilePath);
                }
                else
                {
                    nameElement.Click(); 
                     
                    string price = "";
                    string transportFee = "";
                    string monthSellCount = "";
                    string district = "";
                    string commentCount = "";

                    AndroidElement mainElement = appAccess.GetElementByXPath("//android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.RelativeLayout/android.widget.RelativeLayout/android.widget.FrameLayout/android.widget.ScrollView/android.widget.LinearLayout/android.widget.ListView", true);

                    //价格
                    AndroidElement priceElement = appAccess.GetElementByClassNameAndText("android.widget.TextView", new string[] { "¥" }, false, true);
                    price = priceElement.Text.Trim();

                    //快递费、月销量、地区
                    XmlElement xmlRoot = appAccess.GetXmlRootElement();
                    XmlNodeList linearLayoutXmlElements = xmlRoot.SelectNodes("//android.widget.ListView[@resource-id=\"com.taobao.taobao:id/mainpage\"]/android.widget.LinearLayout");
                    XmlNodeList otherInfoXmlElements = null;
                    foreach (XmlElement linearLayoutxmlElement in linearLayoutXmlElements)
                    {
                       XmlNodeList checkXmlElements = linearLayoutxmlElement.SelectNodes("./android.widget.TextView");
                       foreach (XmlElement checkXmlElement in checkXmlElements)
                        {
                            string checkText = checkXmlElement.GetAttribute("text");
                            if (checkText.Contains("快递:"))
                            {
                                otherInfoXmlElements = checkXmlElements;
                                break;
                            }
                        }
                       if (otherInfoXmlElements != null)
                        {
                            break;
                        }
                    }
                    if (otherInfoXmlElements != null)
                    {
                        foreach (XmlElement element in otherInfoXmlElements)
                        {
                            string text = element.GetAttribute("text").Trim();
                            if (text.Contains("快递"))
                            {
                                transportFee = text;
                            }
                            else if (text.Contains("月销"))
                            {
                                monthSellCount = text;
                            }
                            else
                            {
                                district = text;
                            }
                        }
                    }

                    //滚动屏幕，用来显示评论数
                    //appAccess.Swipe(new Point(100, 1000), new Point(100, 100), 2000);

                    try
                    {
                        AndroidElement commentElement = appAccess.GetElementById("com.taobao.taobao:id/detail_main_comment_count", true);
                        commentCount = commentElement.Text.Trim();
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("没有找到评论数元素, shopName = " + shopName + ", shopWebPageUrl = " + shopWebPageUrl + ", skuName = " + skuName, ex);
                    }

                    skuInfo = new Dictionary<string, string>();
                    skuInfo.Add("name", skuName);
                    skuInfo.Add("price", price);
                    skuInfo.Add("transportFee", transportFee);
                    skuInfo.Add("monthSellCount", monthSellCount);
                    skuInfo.Add("district", district);
                    skuInfo.Add("commentCount", commentCount);
                    this.SaveSkuInfoToLocalFile(skuFilePath, skuInfo);

                    appAccess.ClickBackButton();
                }
                return skuInfo;
            }
            else
            {
                return null;
            }
        } 

        private Dictionary<string, string> ReadSkuInfoFromLocalFile(string filePath)
        {
            CsvReader cr = new CsvReader(filePath);
            Dictionary<string, string> row = cr.GetFieldValues(0);
            return row;
        }

        private void SaveSkuInfoToLocalFile(string filePath, Dictionary<string, string> skuInfo)
        {
            Dictionary<string, int> columnNameToIndex = new Dictionary<string, int>();
            columnNameToIndex.Add("name", 0);
            columnNameToIndex.Add("price", 1);
            columnNameToIndex.Add("transportFee", 2);
            columnNameToIndex.Add("monthSellCount", 3);
            columnNameToIndex.Add("district", 4);
            columnNameToIndex.Add("commentCount", 5);
            CsvWriter cw = new CsvWriter(filePath, columnNameToIndex);
            cw.AddRow(skuInfo);
            cw.SaveToDisk();
        }

        private string GetSkuInfoLocalFilePath(string shopName, string skuName)
        {
            string skuFilePath = Path.Combine(Path.Combine(this.ExportDir, CommonUtil.ProcessFileName(shopName, "_")), CommonUtil.ProcessFileName(skuName, "_") + ".csv");
            return skuFilePath;
        }

        private List<Dictionary<string, string>> ReadShopFromLocalFile(string filePath)
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

        private void SaveShopToLocalFile(string filePath, List<Dictionary<string, string>> shopSkuInfos)
        {
            Dictionary<string, int> columnNameToIndex = new Dictionary<string, int>();
            columnNameToIndex.Add("name", 0);
            columnNameToIndex.Add("price", 1);
            columnNameToIndex.Add("transportFee", 2);
            columnNameToIndex.Add("monthSellCount", 3);
            columnNameToIndex.Add("district", 4);
            columnNameToIndex.Add("commentCount", 5);
            CsvWriter cw = new CsvWriter(filePath, columnNameToIndex);
            foreach (Dictionary<string, string> skuInfo in shopSkuInfos)
            {
                cw.AddRow(skuInfo);
            }
            cw.SaveToDisk();
        }

        private string GetShopLocalFilePath(string shopName,string shopWebPageUrl)
        {
            string shopFilePath = Path.Combine(this.ExportDir, CommonUtil.ProcessFileName(shopWebPageUrl + "_" + shopName, "_") + ".csv");
            return shopFilePath;
        }

        private void GotoShopSkuPage(AndroidAppAccess appAccess, string shopWebPageUrl, string shopName)
        { 
            //点击后进入查询页面
            AndroidElement gotoInputElement = appAccess.GetElementById("com.taobao.taobao:id/home_searchedit", true);
            gotoInputElement.Click();

            //选择店铺选项卡
            AndroidElement shopTabElement = appAccess.GetElementById("com.taobao.taobao:id/search_tab_layout", true);
            shopTabElement.FindElementsByClassName("android.support.v7.app.ActionBar$Tab")[2].Click();

            //录入店铺名称
            AndroidElement inputShopUrlElement = appAccess.GetElementById("com.taobao.taobao:id/searchEdit", true);
            inputShopUrlElement.SendKeys(shopName);

            //查询按钮，点击后应该进入店铺页面
            AndroidElement searchShopBtnElement = appAccess.GetElementById("com.taobao.taobao:id/searchbtn", true);
            searchShopBtnElement.Click();

            try
            {

                ReadOnlyCollection<AndroidElement> shopListViewElements = appAccess.GetElementsById("com.taobao.taobao:id/shopTitle", true);

                if (shopListViewElements.Count > 0)
                {
                    AndroidElement shopElement = shopListViewElements[0];
                    if (shopElement.Text.Trim() == shopName)
                    {
                        shopElement.Click();

                        try
                        {
                            AndroidElement mainMenuTabContainerElement = appAccess.GetElementById("com.taobao.taobao:id/tl_tabs", true);
                            mainMenuTabContainerElement.FindElementsByClassName("android.support.v7.app.ActionBar$Tab")[1].Click();
                        }
                        catch (Exception ex)
                        {
                            throw new Exception("未成功获取'全部宝贝'按钮, shopName = " + shopName + ", shopUrl = " + shopWebPageUrl);
                        }
                    }
                    else
                    {
                        throw new Exception("没有搜索到此店铺，第一个匹配项不是此店铺. shopName = " + shopName);
                    }
                }
                else
                {
                    throw new Exception("没有搜索到此店铺，关键字匹配0个店铺. shopName = " + shopName);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("没有搜索到此店铺. shopName = " + shopName, ex);
            }
        }
    }
}