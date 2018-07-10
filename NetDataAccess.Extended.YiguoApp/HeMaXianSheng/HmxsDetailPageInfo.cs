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
using System.Xml; 

namespace NetDataAccess.Extended.YiguoApp
{
    /// <summary>
    /// 盒马鲜生
    /// </summary>
    public class HmxsDetailPageInfo : CustomProgramBase
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
            initParams.Add("appPackage", "com.wudaokou.hippo");
            initParams.Add("appActivity", "com.wudaokou.hippo.activity.main.MainActivity");
            //initParams.Add("appActivity", "com.wudaokou.hippo.activity.splash.SplashActivity");
            initParams.Add("url", "http://127.0.0.1:4723/wd/hub");
            AppAccess.InitDriver(initParams);
            Thread.Sleep(5000);
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

        #region GrabData
        private bool GrabData(string parameters)
        {  
            try
            {
                string[] strs = parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);

                //输出位置文件夹
                _ExportDir = strs[0]; 

                InitAppAccess();

                ToSelectLocationPage("联邦快递");

                ToCategoryPage();

                GetAllShopGoodsInfo();

                GenerateAllGoodsFile();

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

        #region GenerateAllGoodsFile
        private bool GenerateAllGoodsFile()
        {
            bool succeed = true;
            List<NcpAppElement> allCategory3Elements = GetNeedGrabCategoryInfoFromFile();

            foreach (NcpAppElement c3Element in allCategory3Elements)
            {
                string goodsFilePath = GetCategoryGoodsFilePath(c3Element.Id);
                if (!File.Exists(goodsFilePath))
                {   
                    succeed = false;
                }
            }
            if (succeed)
            {
                try
                {
                    string[] allGoodsColumns = new string[]{ 
                "商品名称",
                "价格",
                "计量单位",
                "一级分类", 
                "二级分类", 
                "三级分类"};
                    Dictionary<string, int> allGoodsColumnDic = CommonUtil.InitStringIndexDic(allGoodsColumns);

                    Dictionary<string, string> allGoodsColumnFormats = new Dictionary<string, string>();
                    allGoodsColumnFormats.Add("价格", "#0.00");

                    string allGoodsFilePath = Path.Combine(this.ExportDir, "AllGoods.xlsx");
                    ExcelWriter allGoodsEW = new ExcelWriter(allGoodsFilePath, "List", allGoodsColumnDic, allGoodsColumnFormats);

                    foreach (NcpAppElement c3Element in allCategory3Elements)
                    {
                        string goodsFilePath = GetCategoryGoodsFilePath(c3Element.Id); 

                        ExcelReader er = new ExcelReader(goodsFilePath, "List");
                        int rowCount = er.GetRowCount();
                        for (int i = 0; i < rowCount; i++)
                        {
                            Dictionary<string, string> row = er.GetFieldValues(i); 

                            Dictionary<string, object> newRow = new Dictionary<string, object>();
                            newRow.Add("商品名称", row["name"]);
                            newRow.Add("价格", decimal.Parse(row["price"]));
                            newRow.Add("计量单位", row["unit"]);
                            newRow.Add("一级分类", row["category1Name"]);
                            newRow.Add("二级分类", row["category2Name"]);
                            newRow.Add("三级分类", row["category3Name"]); 
                            allGoodsEW.AddRow(newRow);
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
          
        #region 点击退回按钮
        private void ClickBackButton()
        {
            // click back button 
            AppAccess.Driver.KeyEvent(4);
        }
        #endregion

        #region 打开店铺页面获取分类信息
        private void ToCategoryPage()
        {
            //关闭分类菜单页
            AndroidElement categoryMenuElement = AppAccess.GetElementByClassNameAndText("android.widget.TextView", new string[] { "生鲜分类" }, true, true, true);
            categoryMenuElement.Click();

            //获取店铺的商品分类
            GetCategoryInfo(); 
        }
        #endregion

        #region 从文件中获取需要爬取数据的分类的信息
        private List<NcpAppElement> GetNeedGrabCategoryInfoFromFile()
        {
            string categoryFilePath = this.GetCategoryFilePath();
            ExcelReader er = new ExcelReader(categoryFilePath, "List");
            int rowCount = er.GetRowCount();
            List<NcpAppElement> allCategory3Elements = new List<NcpAppElement>();
            for (int i = 0; i < rowCount; i++)
            {
                Dictionary<string, string> row = er.GetFieldValues(i);
                string category1Name = row["category1Name"];
                string category2Name = row["category2Name"];
                string category3Name = row["category3Name"]; 
                NcpAppElement element = new NcpAppElement();
                element.Id = category1Name + "_" + category2Name + "_" + category3Name;
                element.Name = category3Name;
                element.Attributes.Add("category2Name", category2Name);
                element.Attributes.Add("category1Name", category1Name);
                allCategory3Elements.Add(element);
            }
            return allCategory3Elements;
        }
        #endregion

        #region 打开店铺页面获取所有商品信息
        private void GetAllShopGoodsInfo()
        {
            List<NcpAppElement> needGrabCategoryElements = GetNeedGrabCategoryInfoFromFile();
            
            //获取店铺的所有分类的商品信息
            GetAllCategoryGoodsInfo();
        }
        #endregion

        #region 获取店铺的所有分类的商品信息
        private void GetAllCategoryGoodsInfo()
        {
            try
            {
                List<NcpAppElement> allCategory3Elements = GetNeedGrabCategoryInfoFromFile();

                for (int i = 0; i < allCategory3Elements.Count; i++)
                {
                    NcpAppElement category3Element = allCategory3Elements[i];
                    string goodsFilePath = this.GetCategoryGoodsFilePath(category3Element.Id);
                    if (!File.Exists(goodsFilePath))
                    {
                        SelectCategory(category3Element);

                        GetAllGoodsInCategory(category3Element);

                        ClickBackButton();
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
        private void GetAllGoodsInCategory(NcpAppElement category3Element)
        {
            try
            {
                hasNoneNewElementTime = 0;
                GoodsElementList = new NcpAppElementList();
                Size winSize = AppAccess.GetWindowSize();
                AppAccess.SwipeDisplayElements(new Point(300, 1000),
                    new Point(300, 300),
                    1000,
                    4000,
                    GetCategoryGoodsItems);  

                //保存到文件 
                string[] goodsColumns = new string[]{"name",
                "price",
                "unit",
                "category1Name",
                "category2Name",
                "category3Name"};
                Dictionary<string, int> goodsColumnDic = CommonUtil.InitStringIndexDic(goodsColumns);
                string goodsFilePath = GetCategoryGoodsFilePath(category3Element.Id);
                ExcelWriter goodsEW = new ExcelWriter(goodsFilePath, "List", goodsColumnDic, null);

                string category1Name = category3Element.Attributes["category1Name"];
                string category2Name = category3Element.Attributes["category2Name"];
                string category3Name = category3Element.Name;

                for (int i = 0; i < GoodsElementList.Count; i++)
                {
                    NcpAppElement element = GoodsElementList[i];
                    string name = element.Name;
                    string price = element.Attributes["price"];
                    string unit = element.Attributes["unit"];
                    Dictionary<string, string> row = new Dictionary<string, string>();
                    row.Add("name", name);
                    row.Add("price", price);
                    row.Add("unit", unit);
                    row.Add("category1Name", category1Name);
                    row.Add("category2Name", category2Name);
                    row.Add("category3Name", category3Name); 
                    goodsEW.AddRow(row); 
                }
                goodsEW.SaveToDisk();

            }
            catch (Exception ex)
            {
                throw new Exception("获取商品信息失败, categoryId = " + category3Element.Id, ex);
            }
        }
        private int hasNoneNewElementTime = 0;
        private bool GetCategoryGoodsItems(int pageIndex)
        { 
            bool hasNewElement = false;
            List<XmlNode> allGoodsParentNodes = AppAccess.GetXmlElementsByXPath(new string[]{
                "//android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.FrameLayout/android.widget.FrameLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.GridView/android.widget.FrameLayout",
                "//android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.RelativeLayout/android.widget.ListView/android.widget.GridView/android.widget.FrameLayout",
                "//android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.RelativeLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.ListView/android.widget.GridView/android.widget.FrameLayout"},
                false);
            if (allGoodsParentNodes != null)
            {
                foreach (XmlNode goodsParentNode in allGoodsParentNodes)
                {
                    XmlNodeList goodsInfoNodes = goodsParentNode.SelectNodes("./android.widget.RelativeLayout/android.widget.TextView");
                    if (goodsInfoNodes.Count >= 3)
                    {
                        string nameValue = goodsInfoNodes[goodsInfoNodes.Count - 3].Attributes["text"].Value;
                        string priceValue = goodsInfoNodes[goodsInfoNodes.Count - 2].Attributes["text"].Value;
                        string unitValue = goodsInfoNodes[goodsInfoNodes.Count - 1].Attributes["text"].Value;
                        if (!GoodsElementList.Exist(nameValue) && priceValue.StartsWith("¥"))
                        {
                            NcpAppElement appElement = GoodsElementList.Add(nameValue, nameValue, "");
                            decimal price = decimal.Parse(priceValue.Substring(1).Trim());
                            appElement.Attributes.Add("price", price.ToString());
                            appElement.Attributes.Add("unit", unitValue.Substring(1));
                            hasNewElement = true;
                        }
                    }
                }
            }
            if (!hasNewElement && hasNoneNewElementTime < 2)
            {
                hasNoneNewElementTime++;
                hasNewElement = true;
                Thread.Sleep(2000);
            }
            else
            {
                hasNoneNewElementTime = 0;
            }
            return hasNewElement;
        }
        #endregion

        #region 选择一个分类，并点击显示其商品
        private void SelectCategory(NcpAppElement categoryElement)
        {
            string category1Name = categoryElement.Attributes["category1Name"];
            AndroidElement c1Node = AppAccess.GetElementByXPath("//android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.RelativeLayout/android.widget.LinearLayout[1]/android.widget.ListView/android.widget.RelativeLayout/android.widget.RelativeLayout/android.widget.TextView[1]", new string[] { category1Name }, true);
            c1Node.Click();

            Thread.Sleep(1500);

            List<AppiumWebElement> cElements = this.GetCategoryElement(categoryElement.Name);
            AppiumWebElement cElement = cElements[cElements.Count - 1];
            Point point = cElement.Location;
            Size size = cElement.Size;
            //点击文字上方的图标
            AppAccess.Tap(1, point.X + size.Width / 2, point.Y - 80, 100);
            //cagtegoryElement.Click();
        }
        #endregion 
        
        #region 打开选择店铺页面
        private void ToSelectLocationPage(string locationName)
        {
            try
            {
                AndroidElement locationClickElement = AppAccess.GetElementByClassNameAndText("android.widget.TextView", new string[] { "选择配送地址" }, true, true, true);
                locationClickElement.Click();

                Thread.Sleep(2000);

                AndroidElement locationInputElement = AppAccess.GetElementByClassNameAndText("android.widget.EditText", new string[] { "输入地址关键字" }, true, true, true);
                locationInputElement.Click();
                locationInputElement.SendKeys(locationName);

                Thread.Sleep(4000);

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
        private string GetCategoryFilePath()
        {
            string categoryFilePath =Path.Combine(ExportDir, "Category.xlsx");
            return categoryFilePath;
        }
        #endregion

        #region 获取某个分类中商品信息文件地址
        private string GetCategoryGoodsFilePath(string categoryFullName)
        {
            string categoryGoodsFilePath = Path.Combine(Path.Combine(ExportDir, "Goods"), categoryFullName + ".xlsx");
            return categoryGoodsFilePath;
        }
        #endregion

        #region 获取分类信息
        private void GetCategoryInfo()
        {
            string categoryFilePath = this.GetCategoryFilePath();

            if (!File.Exists(categoryFilePath))
            {

                try
                {
                    //获取一级分类信息
                    NcpAppElementList allC1Elements = this.GetCategory1Items();

                    foreach (NcpAppElement c1Element in allC1Elements)
                    {
                        //AndroidElement c1Node = AppAccess.GetElementByXPath("//android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.RelativeLayout/android.widget.LinearLayout[1]/android.widget.ListView/android.widget.RelativeLayout/android.widget.TextView[1]", new string[] { c1Element.Name }, true);
                        AndroidElement c1Node = AppAccess.GetElementByXPath("//android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.RelativeLayout/android.widget.LinearLayout[1]/android.widget.ListView/android.widget.RelativeLayout/android.widget.RelativeLayout/android.widget.TextView[1]", new string[] { c1Element.Name }, true);
                        c1Node.Click();

                        CategoryElementList = new NcpAppElementList();
                        Size winSize = AppAccess.GetWindowSize();
                        AppAccess.SwipeDisplayElements(new Point(300, 1000),
                            new Point(300, 300),
                            2000,
                            2000,
                            GetSubCategoryItems);

                        //分级处理
                        List<NcpAppElement> level2Elements = new List<NcpAppElement>();
                        NcpAppElement lastLevel2Element = null;
                        for (int i = 0; i < CategoryElementList.Count; i++)
                        {
                            NcpAppElement element = CategoryElementList[i];
                            if (element.TypeName == "Level2")
                            {
                                level2Elements.Add(element);
                                lastLevel2Element = element;
                            }
                            else
                            {
                                lastLevel2Element.Children.Add(element);
                            }
                        }
                        c1Element.Children.AddRange(level2Elements);
                    }



                    //保存到文件

                    string[] categoryColumns = new string[]{ "category1Name",
                "category2Name", 
                "category3Name"};
                    Dictionary<string, int> categoryColumnDic = CommonUtil.InitStringIndexDic(categoryColumns);
                    ExcelWriter categoryEW = new ExcelWriter(categoryFilePath, "List", categoryColumnDic, null);

                    for (int i = 0; i < allC1Elements.Count; i++)
                    {
                        NcpAppElement c1Element = allC1Elements[i];
                        string category1Name = c1Element.Name;
                        if (c1Element.Children.Count > 0)
                        {
                            for (int j = 0; j < c1Element.Children.Count; j++)
                            {
                                NcpAppElement c2Element = c1Element.Children[j];
                                string category2Name = c2Element.Name;
                                if (c2Element.Children.Count > 0)
                                {
                                    for (int k = 0; k < c2Element.Children.Count; k++)
                                    {
                                        NcpAppElement c3Element = c2Element.Children[k];
                                        string category3Name = c3Element.Name;

                                        Dictionary<string, string> row = new Dictionary<string, string>();
                                        row.Add("category1Name", category1Name);
                                        row.Add("category2Name", category2Name);
                                        row.Add("category3Name", category3Name);
                                        categoryEW.AddRow(row);
                                    }
                                }
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

        private NcpAppElementList GetCategory1Items()
        {
            NcpAppElementList allC1Elements = new NcpAppElementList();
            //ReadOnlyCollection<AndroidElement> allC1Nodes = AppAccess.GetElementsByXPath("//android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.RelativeLayout/android.widget.LinearLayout[1]/android.widget.ListView/android.widget.RelativeLayout/android.widget.TextView[1]", true);
            ReadOnlyCollection<AndroidElement> allC1Nodes = AppAccess.GetElementsByXPath("//android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.RelativeLayout/android.widget.LinearLayout[1]/android.widget.ListView/android.widget.RelativeLayout/android.widget.RelativeLayout/android.widget.TextView[1]", true);
            foreach (AppiumWebElement c1Node in allC1Nodes)
            {
                NcpAppElement c1Element = new NcpAppElement();
                c1Element.Name = c1Node.Text;
                c1Element.Id = c1Element.Name;
                allC1Elements.Add(c1Element);
            }
            return allC1Elements;
        }

        private bool GetSubCategoryItems(int pageIndex)
        {
            int pageOffset = pageIndex * 2000;
            bool hasNewElement = false;
            XmlNodeList allC2ParentNodes = AppAccess.GetXmlElementsByXPath("//android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.RelativeLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.ListView/android.widget.LinearLayout", true);
            foreach (XmlNode c2ParentNode in allC2ParentNodes)
            {
                XmlNode c2Node = c2ParentNode.SelectSingleNode("./android.widget.LinearLayout/android.widget.TextView");
                if (c2Node != null)
                {
                    string text = c2Node.Attributes["text"].Value;
                    string id = text + "Level2";
                    if (!_CategoryElementList.Exist(id))
                    {
                        NcpAppElement appElement = CategoryElementList.Add(id, text, "Level2");
                        hasNewElement = true;
                    }
                }
                XmlNodeList allC3Nodes = c2ParentNode.SelectNodes("./android.widget.LinearLayout/android.widget.GridView/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.TextView");
                if (allC3Nodes != null)
                {
                    foreach (XmlNode c3Node in allC3Nodes)
                    {
                        string text = c3Node.Attributes["text"].Value;
                        string id = text + "Level3";
                        if (!_CategoryElementList.Exist(id))
                        {
                            NcpAppElement appElement = CategoryElementList.Add(id, text, "Level3");
                            hasNewElement = true;
                        }
                    }

                }
            }
            return hasNewElement;
        }
        #endregion

        #region 已确定此界面包含了想要的分类元素，用此函数返回之
        private List<AppiumWebElement> GetCategoryElementAfterFound(string name)
        {
            //出现过找到元素了，但是点击不到的情况，那么再滑屏一下，使得元素暴露的更明显
            try
            {
                AppAccess.Swipe(new Point(300, 600),
                    new Point(300, 10),
                    2000);
            }
            catch (Exception ex)
            {
                //出错，但是界面看起来没问题，所以忽略了
            }
            AndroidElement categoryMainPageElement = AppAccess.GetElementByClassNameAndText("android.widget.TextView", new string[] { "首页" }, true, true, true);
            categoryMainPageElement.Click();

            AndroidElement categoryMenuElement = AppAccess.GetElementByClassNameAndText("android.widget.TextView", new string[] { "生鲜分类" }, true, true, true);
            categoryMenuElement.Click();
            AndroidElement listElement = AppAccess.GetElementByXPath("//android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.RelativeLayout/android.widget.LinearLayout/android.widget.FrameLayout", false);
            //ReadOnlyCollection<AndroidElement> allC3Elements = AppAccess.GetElementsByXPath("//android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.RelativeLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.ListView/android.widget.LinearLayout/android.widget.LinearLayout/android.widget.GridView/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.TextView", false);
            ReadOnlyCollection<AppiumWebElement> allC3Elements = listElement.FindElementsByClassName("android.widget.TextView");
            List<AppiumWebElement> foundElements = new List<AppiumWebElement>();
            foreach (AppiumWebElement element in allC3Elements)
            {
                if (element.Text == name)
                {
                    foundElements.Add(element);
                }
            }
            if (foundElements.Count == 0)
            {
                throw new Exception("GetCategoryElementAfterFound执行出错!");
            }
            else
            {
                return foundElements;
            }
        }
        #endregion

        #region 根据名称和商品数，获取分类元素
        private List<AppiumWebElement> GetCategoryElement(string name)
        {
            AndroidElement categoryMainPageElement = AppAccess.GetElementByClassNameAndText("android.widget.TextView", new string[] { "首页" }, true, true, true);
            categoryMainPageElement.Click();

            AndroidElement categoryMenuElement = AppAccess.GetElementByClassNameAndText("android.widget.TextView", new string[] { "生鲜分类" }, true, true, true);
            categoryMenuElement.Click();
            
            AndroidElement listElement = AppAccess.GetElementByXPath("//android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.RelativeLayout/android.widget.LinearLayout/android.widget.FrameLayout", false);
            //ReadOnlyCollection<AndroidElement> allC3Elements = AppAccess.GetElementsByXPath("//android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.RelativeLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.ListView/android.widget.LinearLayout/android.widget.LinearLayout/android.widget.GridView/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.TextView", false);
            ReadOnlyCollection<AppiumWebElement> allC3Elements = listElement.FindElementsByClassName("android.widget.TextView");
            List<AppiumWebElement> foundElements = new List<AppiumWebElement>();
            for(int i=0;i<allC3Elements.Count;i++)
            {
                AppiumWebElement element = allC3Elements[i];
                if (element.Text == name)
                {
                    foundElements.Add(element);
                    /*
                    if (i + 3 >= allC3Elements.Count)
                    {
                        return GetCategoryElementAfterFound(name);
                    }
                    else
                    {
                        return element;
                    }*/
                }
            }
            if (foundElements.Count > 0)
            {
                return foundElements;
            }

            try
            {
                try
                {
                    Size winSize = AppAccess.GetWindowSize();
                    AppAccess.Swipe(new Point(300, 600),
                        new Point(300, 300),
                        1000);
                }
                catch (Exception ex1)
                {
                }
                Thread.Sleep(1000);
                return this.GetCategoryElement(name);
            }
            catch (Exception ex)
            {
                throw new Exception("找不到分类元素, name = " + name, ex);
            }
        }
        #endregion
    }
}