using System;
using System.Collections.Generic;
using System.Text;
using NetDataAccess.Base.DLL;
using NetDataAccess.Base.Config;
using System.Threading;
using System.Windows.Forms; 
using NetDataAccess.Base.Definition;
using System.IO;
using NetDataAccess.Base.Common;
using NPOI.SS.UserModel;
using NetDataAccess.Base.EnumTypes;
using NetDataAccess.Base.UI; 
using NetDataAccess.Base.Writer;
using NetDataAccess.Base.Reader;
using NetDataAccess.Base.CsvHelper;
using NetDataAccess.Base.DB;
using OpenQA.Selenium.Appium.Android;
using NetDataAccess.AppAccessBase;
using System.Drawing;
using System.Collections.ObjectModel;
using OpenQA.Selenium.Appium; 

namespace NetDataAccess.Extended.YiguoApp
{
    /// <summary>
    /// 京东到家数据
    /// </summary>
    public class JddjDetailPageInfo : CustomProgramBase
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

        #region AndroidAppAccess
        private AndroidAppAccess _AppAccess = null;
        private AndroidAppAccess AppAccess
        {
            get
            {
                return _AppAccess;
            }
        }

        private void CloseAppAccess()
        {
            if (AppAccess != null)
            {
                AppAccess.Close();
            }
        }
        #endregion

        #region 初始化与手机App的连接
        private void InitAppAccess()
        {
            _AppAccess = new AndroidAppAccess();
            Dictionary<string,string> initParams =new Dictionary<string,string>();
            initParams.Add("deviceName", "HUAWEI G700-U00");
            initParams.Add("platformVersion", "4.2");
            initParams.Add("appPackage", "com.jingdong.pdj");
            initParams.Add("appActivity", "pdj.start.StartActivity");
            initParams.Add("url", "http://127.0.0.1:4723/wd/hub");
            AppAccess.InitDriver(initParams);
            Thread.Sleep(10000);
        }
        #endregion

        #region Run
        public bool Run(string parameters, IListSheet listSheet)
        {
            while (!GrabData(parameters))
            { }
            return true;
        }
        #endregion

        #region 定位到城市-上海
        private bool GrabData(string parameters)
        { 
            bool succeed = true;
            try
            {
                string[] strs = parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);

                //输出位置文件夹
                _ExportDir = strs[0];

                //所有地址
                List<string> locationNames = new List<string>();
                for (int i = 1; i < strs.Length; i++)
                {
                    locationNames.Add(strs[i]);
                }

                InitAppAccess();
                GotoShanghai();

                foreach (string locationName in locationNames)
                {

                    ToSelectLocationPage(locationName);

                    GetAllShopCategoryInfo(locationName);

                    GetAllShopGoodsInfo(locationName);

                    GotoSearchLocationPage();

                }
                succeed = GenerateAllGoodsFile(locationNames);

                return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                CloseAppAccess();
            }
        }
        #endregion

        #region 定位到城市-上海
        private bool GenerateAllGoodsFile(List<string> locationNames)
        {
            bool succeed = true;
            foreach (string locationName in locationNames)
            {
                string categoryFilePath = GetCategoryFilePath(locationName);
                if (File.Exists(categoryFilePath))
                {
                    List<NcpAppElement> appElements = GetNeedGrabCategoryInfoFromFile(locationName);
                    foreach (NcpAppElement appElement in appElements)
                    {
                        string goodsFilePath = GetCategoryGoodsFilePath(locationName, appElement.Id);
                        if (!File.Exists(goodsFilePath))
                        {
                            succeed = false;
                        }
                    }
                }
                else
                {
                    succeed = false;
                }
            }
            if (succeed)
            {
                try
                {
                    string[] allGoodsColumns = new string[]{"id",
                "name",
                "price",
                "category1Name", 
                "category2Name", 
                "locationName"};
                    Dictionary<string, int> allGoodsColumnDic = CommonUtil.InitStringIndexDic(allGoodsColumns);

                    Dictionary<string, string> allGoodsColumnFormats = new Dictionary<string, string>();
                    allGoodsColumnFormats.Add("price", "#0.00");

                    string allGoodsFilePath = Path.Combine(this.ExportDir, "AllGoods.xlsx");
                    ExcelWriter allGoodsEW = new ExcelWriter(allGoodsFilePath, "List", allGoodsColumnDic, allGoodsColumnFormats);

                    foreach (string locationName in locationNames)
                    {
                        string categoryFilePath = GetCategoryFilePath(locationName);
                        List<NcpAppElement> appElements = GetNeedGrabCategoryInfoFromFile(locationName);
                        foreach (NcpAppElement appElement in appElements)
                        {
                            string goodsFilePath = GetCategoryGoodsFilePath(locationName, appElement.Id);
                            string parentCategoryName = appElement.Attributes["parentCategory"];

                            ExcelReader er = new ExcelReader(goodsFilePath, "List");
                            int rowCount = er.GetRowCount();
                            for (int i = 0; i < rowCount; i++)
                            {
                                Dictionary<string, string> row = er.GetFieldValues(i);
                                string category1Name = CommonUtil.IsNullOrBlank(parentCategoryName) ? row["categoryName"] : parentCategoryName;
                                string category2Name = CommonUtil.IsNullOrBlank(parentCategoryName) ? "" : row["categoryName"];

                                Dictionary<string, object> newRow = new Dictionary<string, object>();
                                newRow.Add("name", row["name"]);
                                newRow.Add("price", decimal.Parse(row["price"]));
                                newRow.Add("category1Name", category1Name);
                                newRow.Add("category2Name", category2Name);
                                newRow.Add("locationName", locationName);
                                allGoodsEW.AddRow(newRow);
                            }
                        }
                    }
                    allGoodsEW.SaveToDisk();
                    succeed = true;
                }
                catch (Exception ex)
                {
                    throw new Exception("输出全部商品文件出错", ex);
                }
            }
            return succeed;
        }
        #endregion

        #region 定位到城市-上海
        private void GotoShanghai()
        {
            //AndroidElement locationTitleElement = AppAccess.GetElementByClassNameAndText("android.widget.TextView", new string[] { "无法获取地址", "正在获得您的位置" }, true);
            //locationTitleElement.Click();

            GotoSearchLocationPage();

            AndroidElement currentCityElement = AppAccess.GetElementByClassNameAndText("android.widget.TextView", new string[] { "当前城市" }, true, true, true);
            currentCityElement.Click();

            AndroidElement shanghaiElement = AppAccess.GetElementByClassNameAndText("android.widget.TextView", new string[] { "上海市" }, true, true, true);
            shanghaiElement.Click();
        }
        #endregion

        #region 跳转到定位小区区域页面
        private void GotoSearchLocationPage()
        {
            AndroidElement locationTitleElement = AppAccess.GetElementByXPath("//android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.TabHost/android.widget.FrameLayout/android.widget.RelativeLayout/android.widget.LinearLayout/android.widget.RelativeLayout/android.widget.RelativeLayout/android.widget.TextView", true);
            locationTitleElement.Click();
        }
        #endregion

        #region 点击退回按钮
        private void ClickBackButton()
        {
            AndroidElement backElement = AppAccess.GetElementByXPath("//android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.RelativeLayout/android.widget.RelativeLayout/android.widget.ImageView", true);
            backElement.Click();
        }
        #endregion

        #region 打开店铺页面获取分类信息
        private void GetAllShopCategoryInfo(string locationName)
        {
            string categoryFilePath = this.GetCategoryFilePath(locationName);

            if (!File.Exists(categoryFilePath))
            {
                GotoYonghuiShopPage();

                //获取店铺的商品分类
                GetCategoryInfo(locationName);

                //关闭分类菜单页
                AndroidElement categoryMenuElement = AppAccess.GetElementByClassNameAndText("android.widget.TextView", new string[] { "分类" }, true, true, true);
                categoryMenuElement.Click();

                ClickBackButton(); 
            }
        }
        #endregion

        #region 从文件中获取需要爬取数据的分类的信息
        private List<NcpAppElement> GetNeedGrabCategoryInfoFromFile(string locationName)
        {
            string categoryFilePath = this.GetCategoryFilePath(locationName);
            ExcelReader er = new ExcelReader(categoryFilePath, "List");
            int rowCount = er.GetRowCount();
            List<NcpAppElement> allCategoryElements = new List<NcpAppElement>();
            for (int i = 0; i < rowCount; i++)
            {
                Dictionary<string, string> row = er.GetFieldValues(i);
                string id = row["id"];
                string category1Name = row["category1Name"]; 
                string category2Name = row["category2Name"];
                string needGrab = row["needGrab"];
                string goodsCount = row["goodsCount"];
                if (needGrab == "是")
                {
                    string name = CommonUtil.IsNullOrBlank(category2Name) ? category1Name : category2Name;
                    string parentCategory = CommonUtil.IsNullOrBlank(category2Name) ? "" : category1Name;
                    NcpAppElement element = new NcpAppElement();
                    element.Id = id;
                    element.Name = name;
                    element.Attributes.Add("count", goodsCount);
                    element.Attributes.Add("parentCategory", parentCategory);
                    allCategoryElements.Add(element);
                }
            }
            return allCategoryElements;
        }
        #endregion

        #region 打开店铺页面获取所有商品信息
        private void GetAllShopGoodsInfo(string locationName)
        {
            GotoYonghuiShopPage();

            //获取店铺的所有分类的商品信息
            GetAllCategoryGoodsInfo(locationName);

            //退回到首页
            AndroidElement backElement = AppAccess.GetElementByClassNameAndIndex("android.widget.ImageView", 0, true);
            backElement.Click(); 


        }
        #endregion

        #region 获取店铺的所有分类的商品信息
        private void GetAllCategoryGoodsInfo(string locationName)
        {
            try
            {
                List<NcpAppElement> needGrabCategoryElements = GetNeedGrabCategoryInfoFromFile(locationName);

                for (int i = 0; i < needGrabCategoryElements.Count; i++)
                {
                    NcpAppElement categoryElement = needGrabCategoryElements[i];
                    string goodsFilePath = this.GetCategoryGoodsFilePath(locationName, categoryElement.Id);
                    if (!File.Exists(goodsFilePath))
                    {
                        SelectCategory(categoryElement);

                        GetAllGoodsInCategory(locationName, categoryElement);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("获取店铺的所有分类的商品信息失败.", ex);
            }
        }
        #endregion

        #region 获取这个分类下所有商品信息
        private NcpAppElementList _GoodsElementList = null;
        private NcpAppElementList GoodsElementList
        {
            get
            {
                return _GoodsElementList;
            }
            set
            {
                _GoodsElementList = value;
            }
        }
        private void GetAllGoodsInCategory(string locationName, NcpAppElement categoryElement)
        {
            try
            {
                hasNoneNewElementTime = 0;
                GoodsElementList = new NcpAppElementList();
                Size winSize = AppAccess.GetWindowSize();
                AppAccess.SwipeDisplayElements(new Point(winSize.Width - 20, winSize.Height - 200),
                    new Point(winSize.Width - 20, 400),
                    1000,
                    10000,
                    GetCategoryGoodsItems);  

                //保存到文件
                int goodsCount = int.Parse(categoryElement.Attributes["count"]); 
                string[] goodsColumns = new string[]{"id",
                "name",
                "price",
                "categoryName", 
                "locationName"};
                Dictionary<string, int> goodsColumnDic = CommonUtil.InitStringIndexDic(goodsColumns);
                string goodsFilePath = (goodsCount == GoodsElementList.Count ? this.GetCategoryGoodsFilePath(locationName, categoryElement.Id) : this.GetCategoryGoodsFilePath("_Error_" + locationName, categoryElement.Id));
                ExcelWriter goodsEW = new ExcelWriter(goodsFilePath, "List", goodsColumnDic, null);

                for (int i = 0; i < GoodsElementList.Count; i++)
                {
                    NcpAppElement element = GoodsElementList[i];
                    string id = element.Id;
                    string name = element.Name;
                    string price = element.Attributes["price"];  
                    Dictionary<string, string> row = new Dictionary<string, string>();
                    row.Add("id", element.Id);
                    row.Add("name", name);
                    row.Add("categoryName", categoryElement.Name);
                    row.Add("price", price);
                    row.Add("locationName", locationName);
                    goodsEW.AddRow(row); 
                }
                goodsEW.SaveToDisk();

            }
            catch (Exception ex)
            {
                throw new Exception("获取商品信息失败, locationName = " + locationName + ", categoryId = " + categoryElement.Id, ex);
            }
        }
        private int hasNoneNewElementTime = 0;
        private bool GetCategoryGoodsItems(int pageIndex)
        { 
            bool hasNewElement = false;
            ReadOnlyCollection<AndroidElement> allLinearLayouts = AppAccess.GetElementsByXPath("//android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.RelativeLayout/android.view.View/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.RelativeLayout/android.widget.LinearLayout/android.support.v7.widget.RecyclerView/android.widget.LinearLayout", true);
            foreach (AppiumWebElement element in allLinearLayouts)
            {
                ReadOnlyCollection<AppiumWebElement> imageElements = element.FindElementsByClassName("android.widget.ImageView");
                ReadOnlyCollection<AppiumWebElement> textElements = element.FindElementsByClassName("android.widget.TextView");
                if (imageElements != null && imageElements.Count > 0 && textElements != null && textElements.Count >= 2)
                {
                    AppiumWebElement nameElement = textElements[textElements.Count - 2];
                    AppiumWebElement priceElement = textElements[textElements.Count - 1];
                    string id = nameElement.Text + priceElement.Text;
                    if (!GoodsElementList.Exist(id) && !nameElement.Text.StartsWith("￥") && priceElement.Text.StartsWith("￥"))
                    {
                        Point location = new Point(element.Location.X, element.Location.Y);
                        NcpAppElement appElement = GoodsElementList.Add(id, nameElement.Text, "", location, element.Size);
                        appElement.Attributes.Add("price", priceElement.Text.Substring(1));
                        hasNewElement = true;
                    }
                }
            }
            if (!hasNewElement && hasNoneNewElementTime == 0)
            {
                hasNoneNewElementTime++;
                hasNewElement = true;
            }
            return hasNewElement;
        }
        #endregion

        #region 选择一个分类，并点击显示其商品
        private void SelectCategory(NcpAppElement categoryElement)
        {
            AndroidElement categoryMenuElement = AppAccess.GetElementByClassNameAndText("android.widget.TextView", new string[] { "分类" }, true, true, true);
            categoryMenuElement.Click();

            AppiumWebElement cagtegoryElement = this.GetCategoryElement(categoryElement.Name, categoryElement.Attributes["count"]);
            cagtegoryElement.Click();
        }
        #endregion

        #region 打开选择location页面
        private void ToSelectLocationPage(string locationName)
        {
            try
            {
                AndroidElement locationInputElement = AppAccess.GetElementByClassNameAndText("android.widget.EditText", new string[] { "写字楼、小区、学校" }, true, true, true);
                locationInputElement.Click();
                locationInputElement.SendKeys(locationName);

                AndroidElement locationOptionElement = AppAccess.GetElementByClassNameAndText("android.widget.TextView", new string[] { locationName }, true, true, true);
                locationOptionElement.Click(); 
            }
            catch (Exception ex)
            {
                throw new Exception("打开选择location页面, locatoinName = " + locationName, ex);
            }
        }
        #endregion

        #region 打开店铺页面
        private void GotoYonghuiShopPage()
        {
            AndroidElement yonghuiElement = AppAccess.GetElementByClassNameAndText("android.widget.TextView", new string[] { "永辉超市" }, true, true, true);
            yonghuiElement.Click();
        }
        #endregion

        #region 获取分类信息文件地址
        private string GetCategoryFilePath(string locationName)
        {
            string categoryFilePath = Path.Combine(Path.Combine(ExportDir, "Category"), locationName + ".xlsx");
            return categoryFilePath;
        }
        #endregion

        #region 获取某个分类中商品信息文件地址
        private string GetCategoryGoodsFilePath(string locationName, string categoryId)
        {
            string categoryGoodsFilePath = Path.Combine(Path.Combine(ExportDir, "Goods"), locationName + "_" + categoryId + ".xlsx");
            return categoryGoodsFilePath;
        }
        #endregion

        #region 获取分类信息
        private void GetCategoryInfo(string locationName)
        {
            try
            {
                AndroidElement categoryMenuElement = AppAccess.GetElementByClassNameAndText("android.widget.TextView", new string[] { "分类" }, true, true, true);
                categoryMenuElement.Click();

                CategoryElementList = new NcpAppElementList();
                Size winSize = AppAccess.GetWindowSize();
                AppAccess.SwipeDisplayElements(new Point(winSize.Width - 20, winSize.Height - 20), 
                    new Point(winSize.Width - 20, 200), 
                    2000, 
                    5000, 
                    GetCategoryItems);
                NcpAppElementList sortedCategoryElements = CategoryElementList.SortByPosition();

                //分级处理
                List<NcpAppElement> level1Elements = new List<NcpAppElement>();
                NcpAppElement lastLevel1Element = null;
                for (int i = 1; i < sortedCategoryElements.Count; i++)
                {
                    NcpAppElement element = sortedCategoryElements[i];
                    if (element.TypeName == "Level1")
                    {
                        level1Elements.Add(element);
                        lastLevel1Element = element;
                    }
                    else
                    {
                        lastLevel1Element.Children.Add(element);
                    }
                }

                //保存到文件

                string[] categoryColumns = new string[]{"id",
                "category1Name",
                "category2Name", 
                "needGrab",
                "goodsCount", 
                "locationName"};
                Dictionary<string, int> categoryColumnDic = CommonUtil.InitStringIndexDic(categoryColumns);
                string categoryFilePath = this.GetCategoryFilePath(locationName);
                ExcelWriter categoryEW = new ExcelWriter(categoryFilePath, "List", categoryColumnDic, null);

                for (int i = 0; i < level1Elements.Count; i++)
                {
                    NcpAppElement element = level1Elements[i];
                    string category1Name = element.Name;
                    string needGrab = element.Children.Count == 0 ? "是" : "否";
                    string goodsCount = element.Attributes["count"];
                    Dictionary<string, string> row = new Dictionary<string, string>();
                    row.Add("id", element.Id);
                    row.Add("category1Name", category1Name);
                    row.Add("category2Name", "");
                    row.Add("needGrab", needGrab);
                    row.Add("goodsCount", goodsCount);
                    row.Add("locationName", locationName);
                    categoryEW.AddRow(row);
                    if (element.Children.Count > 0)
                    {
                        for (int j = 0; j < element.Children.Count; j++)
                        {
                            NcpAppElement childElement = element.Children[j];
                            string category2Name = childElement.Name;
                            string childGoodsCount = childElement.Attributes["count"];
                            Dictionary<string, string> childRow = new Dictionary<string, string>();
                            childRow.Add("id", childElement.Id);
                            childRow.Add("category1Name", category1Name);
                            childRow.Add("category2Name", category2Name);
                            childRow.Add("needGrab", "是");
                            childRow.Add("goodsCount", childGoodsCount);
                            childRow.Add("locationName", locationName);
                            categoryEW.AddRow(childRow);
                        }
                    }
                }
                categoryEW.SaveToDisk();

            }
            catch (Exception ex)
            {
                throw new Exception("获取分类信息失败.", ex);
            }
        }
        #endregion

        #region 翻页获取分类信息
        private NcpAppElementList _CategoryElementList = null;
        private NcpAppElementList CategoryElementList
        {
            get
            {
                return _CategoryElementList;
            }
            set
            {
                _CategoryElementList = value;
            }
        }

        private bool GetCategoryItems(int pageIndex)
        {
            int pageOffset = pageIndex * 2000;
            bool hasNewElement = false;
            ReadOnlyCollection<AndroidElement> allRelativeLayouts = AppAccess.GetElementsByXPath("//android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.RelativeLayout/android.view.View/android.widget.FrameLayout[2]/android.widget.LinearLayout/android.support.v7.widget.RecyclerView/android.widget.RelativeLayout", true);
            foreach (AppiumWebElement element in allRelativeLayouts)
            {
                ReadOnlyCollection<AppiumWebElement> textElements = element.FindElementsByClassName("android.widget.TextView");
                if (textElements != null && textElements.Count >= 2)
                {
                    //需要改造，将名称+宽度作为唯一标识，因为商品数量经常变
                    //获取某类下的所有商品时，商品数量要按抓取此分类时的数量为准
                    //增加店铺名的抓取
                    AppiumWebElement nameElement = textElements[0];
                    AppiumWebElement countElement = textElements[1];
                    string id = nameElement.Text + countElement.Text;
                    if (!_CategoryElementList.Exist(id))
                    {
                        Point location = new Point(element.Location.X, element.Location.Y + pageOffset);
                        NcpAppElement appElement = CategoryElementList.Add(id, nameElement.Text, "Level1", location, element.Size);
                        appElement.Attributes.Add("count", countElement.Text.Substring(1, countElement.Text.Length - 2));
                        hasNewElement = true;
                    }
                }
            }

            ReadOnlyCollection<AndroidElement> allLinearLayouts = AppAccess.GetElementsByXPath("//android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.RelativeLayout/android.view.View/android.widget.FrameLayout[2]/android.widget.LinearLayout/android.support.v7.widget.RecyclerView/android.widget.LinearLayout", true);
            foreach (AppiumWebElement element in allLinearLayouts)
            {
                ReadOnlyCollection<AppiumWebElement> textElements = element.FindElementsByClassName("android.widget.TextView");
                if (textElements != null && textElements.Count >= 2)
                {
                    AppiumWebElement nameElement = textElements[0];
                    AppiumWebElement countElement = textElements[1];
                    string id = nameElement.Text + countElement.Text;
                    if (!_CategoryElementList.Exist(id))
                    {
                        Point location = new Point(element.Location.X, element.Location.Y + pageOffset);
                        NcpAppElement appElement = CategoryElementList.Add(id, nameElement.Text, "Level2", location, element.Size);
                        appElement.Attributes.Add("count", countElement.Text.Substring(1, countElement.Text.Length - 2));
                        hasNewElement = true;
                    }
                }
            }

            return hasNewElement;
        }
        #endregion

        #region 根据名称和商品数，获取分类元素
        private AppiumWebElement GetCategoryElement(string name, string count)
        {  
            ReadOnlyCollection<AndroidElement> allRelativeLayouts = AppAccess.GetElementsByXPath("//android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.RelativeLayout/android.view.View/android.widget.FrameLayout[2]/android.widget.LinearLayout/android.support.v7.widget.RecyclerView/android.widget.RelativeLayout", true);
            foreach (AppiumWebElement element in allRelativeLayouts)
            {
                ReadOnlyCollection<AppiumWebElement> textElements = element.FindElementsByClassName("android.widget.TextView");
                if (textElements != null && textElements.Count >= 2)
                {
                    AppiumWebElement nameElement = textElements[0];
                    AppiumWebElement countElement = textElements[1];
                    if(nameElement.Text == name && countElement.Text=="("+count+")")
                    {
                        return element;
                    }
                }
            }
             
            ReadOnlyCollection<AndroidElement> allLinearLayouts = AppAccess.GetElementsByXPath("//android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.RelativeLayout/android.view.View/android.widget.FrameLayout[2]/android.widget.LinearLayout/android.support.v7.widget.RecyclerView/android.widget.LinearLayout", true);
            foreach (AppiumWebElement element in allLinearLayouts)
            {
                ReadOnlyCollection<AppiumWebElement> textElements = element.FindElementsByClassName("android.widget.TextView");
                if (textElements != null && textElements.Count >= 2)
                {
                    AppiumWebElement nameElement = textElements[0];
                    AppiumWebElement countElement = textElements[1];
                    if(nameElement.Text == name && countElement.Text=="("+count+")")
                    {
                        return element;
                    }
                }
            }

            try
            {
                Size winSize = AppAccess.GetWindowSize();
                AppAccess.Swipe(new Point(winSize.Width - 20, winSize.Height - 20),
                    new Point(winSize.Width - 20, 200),
                    2000);
                Thread.Sleep(3000);
                return this.GetCategoryElement(name, count);
            }
            catch (Exception ex)
            {
                throw new Exception("找不到分类元素, name = " + name + ", count = " + count, ex);
            }
        }
        #endregion
    }
}