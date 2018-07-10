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
    /// GetAllListPage
    /// </summary>
    [System.Runtime.InteropServices.ComVisibleAttribute(true)]
    public class GetShopList : ExternalRunWebPage
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

        public override void GetDataByOtherAcessType(Dictionary<string, string> listRow)
        {
            string keyword = listRow["keyword"]; 
            string keywordShopFilePath = this.GetKeywordShopFilePath(keyword);
            if (!File.Exists(keywordShopFilePath))
            {

                this.RunPage.InvokeAppendLogText("开始连接手机APP", LogLevelType.System, true);
                AndroidAppAccess appAccess = this.InitAppAccess();
                this.RunPage.InvokeAppendLogText("连接手机APP成功", LogLevelType.System, true);

                try
                {
                    bool hasShop = this.GotoKeywordShopListPage(appAccess, keyword);
                    List<Dictionary<String, string>> allKeywordShops = new List<Dictionary<string, string>>();
                    Dictionary<string, string> allShopNames = new Dictionary<string, string>();
                    if (hasShop)
                    {
                        this.GetKeywordShopInfos(appAccess, keyword, allKeywordShops, allShopNames, 0);
                    }
                    this.SaveKeywordShopInfoToLocalFile(keywordShopFilePath, allKeywordShops, keyword);
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
            string filePath = Path.Combine(this.RunPage.GetExportDir(), "京东店铺商品.xlsx");
            Dictionary<string, int> columnNameToIndex = CommonUtil.InitStringIndexDic(new string[]{
                "detailPageUrl",
                "detailPageName", 
                "cookie",  
                "grabStatus",   
                "giveUpGrab",  
                "name",  
                "subscribe",  
                "mark",  
                "jdSelf",  
                "keyword"});

            ExcelWriter ew = new ExcelWriter(filePath, "List", columnNameToIndex);

            int keywordCount = listSheet.RowCount;
            Dictionary<string, string> nameKeywordDic = new Dictionary<string, string>();
            for (int i = 0; i < keywordCount; i++)
            {
                Dictionary<string, string> listRow = listSheet.GetRow(i);
                string keyword = listRow["keyword"];
                string keywordShopFilePath = this.GetKeywordShopFilePath(keyword);
                List<Dictionary<string, string>> shopInfos = this.ReadKeywordShopFromLocalFile(keywordShopFilePath);
                foreach (Dictionary<string, string> shopInfo in shopInfos)
                {
                    string name = shopInfo["name"]; 
                    string nameKeyword = name + "_" + keyword;
                    if (!nameKeywordDic.ContainsKey(nameKeyword))
                    {
                        nameKeywordDic.Add(name, null);
                        shopInfo.Add("detailPageUrl", name);
                        shopInfo.Add("detailPageName", nameKeyword);
                        ew.AddRow(shopInfo);
                    }
                }
            }
            ew.SaveToDisk(); 

            return base.AfterAllGrab(listSheet);
        }

        private void GetKeywordShopInfos(AndroidAppAccess appAccess, string keyword, List<Dictionary<String, string>> allKeywordShops, Dictionary<string, string> allShopNames, int tryGotTime)
        {
            try
            {

                ReadOnlyCollection<AndroidElement> allShopElements = appAccess.GetElementsById("com.jingdong.app.mall:id/b6j", true);
                XmlElement rootElement = appAccess.GetXmlRootElement();
                XmlNodeList shopXmlElements = rootElement.SelectNodes("//android.widget.LinearLayout[@resource-id=\"com.jingdong.app.mall:id/b6j\"]");
                bool gotNewShop = false;
                for (int i = 0; i < allShopElements.Count; i++)
                {
                    AndroidElement shopElement = allShopElements[i];
                    XmlElement shopXmlElement = (XmlElement)shopXmlElements[i];
                    if (this.GetShopInfo(appAccess, shopElement, shopXmlElement, allKeywordShops, allShopNames))
                    {
                        gotNewShop = true;
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
                if (gotNewShop)
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
                    this.GetKeywordShopInfos(appAccess, keyword, allKeywordShops, allShopNames, tryGotTime);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private bool GetShopInfo(AndroidAppAccess appAccess, AndroidElement shopElement, XmlElement shopXmlElement, List<Dictionary<String, string>> allKeywordShops, Dictionary<string, string> allShopNames)
        {
            try
            {
                XmlElement nameXmlElement = (XmlElement)shopXmlElement.SelectSingleNode("./descendant::android.widget.TextView[@resource-id=\"com.jingdong.app.mall:id/b6n\"]");
                if (nameXmlElement != null)
                {
                    string name = nameXmlElement.GetAttribute("text");
                    if (!allShopNames.ContainsKey(name))
                    {
                        XmlElement subscribeXmlElement = (XmlElement)shopXmlElement.SelectSingleNode("./descendant::android.widget.TextView[@resource-id=\"com.jingdong.app.mall:id/b6s\"]");
                        if (subscribeXmlElement != null)
                        {
                            string subscribe = subscribeXmlElement.GetAttribute("text");
                            XmlElement markXmlElement = (XmlElement)shopXmlElement.SelectSingleNode("./descendant::android.widget.TextView[@resource-id=\"com.jingdong.app.mall:id/b6t\"]");
                            string mark = markXmlElement == null ? "" : markXmlElement.GetAttribute("text");

                            XmlElement jdSelfXmlElement = (XmlElement)shopXmlElement.SelectSingleNode("./descendant::android.widget.TextView[@resource-id=\"com.jingdong.app.mall:id/b6k\"]");
                            string jdSelf = jdSelfXmlElement == null ? "入驻" : "自营";

                            allShopNames.Add(name, null);
                            Dictionary<string, string> shopInfo = new Dictionary<string, string>();
                            shopInfo.Add("name", name);
                            shopInfo.Add("subscribe", subscribe);
                            shopInfo.Add("mark", mark);
                            shopInfo.Add("jdSelf", jdSelf);
                            allKeywordShops.Add(shopInfo);
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

        private void SaveKeywordShopInfoToLocalFile(string filePath, List<Dictionary<string, string>> allKeywordShops, string keyword)
        {
            Dictionary<string, int> columnNameToIndex = new Dictionary<string, int>();
            columnNameToIndex.Add("name", 0);
            columnNameToIndex.Add("subscribe", 1);
            columnNameToIndex.Add("mark", 2);
            columnNameToIndex.Add("jdSelf", 3);
            columnNameToIndex.Add("keyword", 4); 
            CsvWriter cw = new CsvWriter(filePath, columnNameToIndex);
            foreach (Dictionary<string, string> skuInfo in allKeywordShops)
            {
                skuInfo.Add("keyword", keyword);
                cw.AddRow(skuInfo);
            }
            cw.SaveToDisk();
        }

        private string GetKeywordShopFilePath(string keyword)
        {
            string filePath = Path.Combine(this.RunPage.GetExportDir(), CommonUtil.ProcessFileName(keyword, "_") + ".csv");
            return filePath;
        }

        private bool GotoKeywordShopListPage(AndroidAppAccess appAccess, string keyword)
        {
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

            //录入店铺keyword
            AndroidElement inputKeywordElement = appAccess.GetElementById("com.jd.lib.search:id/search_text", true);
            inputKeywordElement.SendKeys(keyword);

            //店铺查询按钮，点击后应该进入店铺列表页面
            AndroidElement searchShopBtnElement = appAccess.GetElementById("com.jd.lib.search:id/home_auto_complete_item_container", true);
            searchShopBtnElement.Click();

            AndroidElement shopListContainerElement = appAccess.GetElementByIdNoWaiting("com.jingdong.app.mall:id/b6b", false);
            int noAlertTime = 0;
            while (shopListContainerElement == null)
            {
                AndroidElement reSearchShopBtnElement = appAccess.GetElementById("com.jingdong.app.mall:id/ad0", true);
                reSearchShopBtnElement.Click();
                Thread.Sleep(3000);
                noAlertTime++;
                shopListContainerElement = appAccess.GetElementByIdNoWaiting("com.jingdong.app.mall:id/b6b", false);
                if (noAlertTime > 3)
                {
                    break;
                }
            }

            if (shopListContainerElement == null)
            {
                AndroidElement noShopAlertElement = appAccess.GetElementByIdNoWaiting("com.jingdong.app.mall:id/b4x", false);
                if (noShopAlertElement != null)
                {
                    return false;
                }
                else
                {
                    throw new Exception("无法获取到店铺列表, keyword = " + keyword);
                }
            }
            else
            {
                return true;
            }
        }
    }
}