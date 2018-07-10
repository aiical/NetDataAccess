using System;
using System.Collections.Generic;
using System.Text;
using OpenQA.Selenium.Appium.Android;
using OpenQA.Selenium.Remote;
using System.Collections.ObjectModel;
using System.Threading;
using NetDataAccess.Base.Common;
using OpenQA.Selenium.Appium;
using System.Drawing;
using OpenQA.Selenium.Appium.Interfaces;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Appium.MultiTouch;
using System.Xml;
using OpenQA.Selenium;

namespace NetDataAccess.AppAccessBase
{
    /// <summary>
    /// 获取Android App数据的通道
    /// </summary>
    public class AndroidAppAccess
    {
        #region AndroidDriver
        private AndroidDriver<AndroidElement> _Driver = null;
        public AndroidDriver<AndroidElement> Driver
        {
            get
            {
                return _Driver;
            }
        }
        #endregion

        #region 初始化AndroidDriver
        public void InitDriver(Dictionary<string,string> initParams)
        {
            string deviceName = "";
            string platformVersion = "";
            string appPackage = "";
            string appActivity = "";
            string url = "";

            if (initParams.ContainsKey("deviceName"))
            {
                deviceName = initParams["deviceName"];
            }
            else
            {
                throw new Exception("初始化AndroidDriver失败, 缺少参数deviceName");
            }
            if (initParams.ContainsKey("platformVersion"))
            {
                platformVersion = initParams["platformVersion"];
            }
            else
            {
                throw new Exception("初始化AndroidDriver失败, 缺少参数platformVersion");
            }
            if (initParams.ContainsKey("appPackage"))
            {
                appPackage = initParams["appPackage"];
            }
            else
            {
                throw new Exception("初始化AndroidDriver失败, 缺少参数appPackage");
            }
            if (initParams.ContainsKey("appActivity"))
            {
                appActivity = initParams["appActivity"];
            }
            else
            {
                throw new Exception("初始化AndroidDriver失败, 缺少参数appActivity");
            }
            if (initParams.ContainsKey("url"))
            {
                url = initParams["url"];
            }
            else
            {
                throw new Exception("缺少参数url");
            }

            DesiredCapabilities capabilities = new DesiredCapabilities();
            capabilities.SetCapability("device", "Android");
            capabilities.SetCapability(CapabilityType.Platform, "Windows");
            capabilities.SetCapability("deviceName", deviceName);
            capabilities.SetCapability("platformName", "Android");
            capabilities.SetCapability("platformVersion", platformVersion);
            capabilities.SetCapability("unicodeKeyboard", true);
            capabilities.SetCapability("resetKeyboard", true); 

            capabilities.SetCapability("appPackage", appPackage);
            capabilities.SetCapability("appActivity", appActivity);

            _Driver = new AndroidDriver<AndroidElement>(new Uri(url), capabilities, TimeSpan.FromMilliseconds(60000));
        }
        #endregion

        #region 操作超时设置
        private int _OperateTimeout = 20000;
        public int OperateTimeout 
        {
            get
            {
                return _OperateTimeout;
            }
            set
            {
                _OperateTimeout = value;
            }
        }
        #endregion

        #region 状态检查间隔时间（毫秒）
        private int _StatusCheckInterval = 200;
        public int StatusCheckInterval
        {
            get
            {
                return _StatusCheckInterval;
            }
            set
            {
                _StatusCheckInterval = value;
            }
        }
        #endregion

        #region 获取分辨率大小
        public Size GetWindowSize()
        {
            Size winSize = Driver.Manage().Window.Size;
            return winSize;
        }
        #endregion

        #region 判断Element是否包含的文字内容
        public bool CheckElementContainText(AndroidElement element, string[] checkStrings, bool andCondtion)
        {
            string text = element.Text;
            return this.CheckContainText(text, checkStrings, andCondtion);
        }
        #endregion

        #region 判断xPath对应的控件里是否包含的文字内容
        public bool CheckContainTextByXPath(string xPath, string[] checkStrings, bool andCondtion)
        {
            ReadOnlyCollection<AndroidElement> elements = Driver.FindElementsByXPath(xPath);
            if (elements == null)
            {
                return false;
            }
            else
            {
                foreach (AndroidElement element in elements)
                {
                    string text = element.Text;
                    if(this.CheckContainText(text, checkStrings, andCondtion))
                    {
                        return true;
                    }
                }
                return false;
            }
        }
        #endregion

        #region 判断App当前页面是否包含的文字内容
        public bool CheckCurrentPageContainText(string[] checkStrings, bool andCondtion)
        {
            string text = Driver.PageSource;
            return this.CheckContainText(text, checkStrings, andCondtion);
        }
        #endregion

        #region 根据类型、序号获取元素
        public bool CheckCurrentPageContainText(string[] checkStrings, bool andCondtion, bool errorNone)
        { 
            DateTime startTime = DateTime.Now;

            while ((DateTime.Now - startTime).TotalMilliseconds < OperateTimeout)
            {
                Thread.Sleep(StatusCheckInterval); 
                if (this.CheckCurrentPageContainText(checkStrings, andCondtion))
                {
                    return true;
                }
            }
            if (errorNone)
            {
                string allText = CommonUtil.StringArrayToString(checkStrings, ",");
                throw new Exception("CheckCurrentPageContainText失败, text = " + allText); 
            }
            else
            {
                return false;
            }
        }
        #endregion

        #region 根据类型、序号获取元素
        public AndroidElement GetElementByClassNameAndIndex(AndroidElement parentElement, string className, int index, bool errorNone)
        {
            DateTime startTime = DateTime.Now;

            while ((DateTime.Now - startTime).TotalMilliseconds < OperateTimeout)
            {
                Thread.Sleep(StatusCheckInterval); 
                ReadOnlyCollection<AppiumWebElement> elements = parentElement.FindElementsByClassName(className);
                if (elements != null && elements.Count > index)
                {
                    return (AndroidElement)elements[index];
                }
            }
            if (errorNone)
            {
                throw new Exception("GetElementByClassNameAndIndex获取元素超时, className = " + className + ", index = " + index.ToString());
            }
            else
            {
                return null;
            }
        }
        #endregion

        #region 根据类型、序号获取元素
        public AndroidElement GetElementByClassNameAndIndex(string className, int index, bool errorNone)
        {
            DateTime startTime = DateTime.Now;

            while ((DateTime.Now - startTime).TotalMilliseconds < OperateTimeout)
            {
                Thread.Sleep(StatusCheckInterval); 
                ReadOnlyCollection<AndroidElement> elements = Driver.FindElementsByClassName(className);
                if (elements != null && elements.Count > index)
                {
                    return elements[index];
                }
            }
            if (errorNone)
            {
                throw new Exception("GetElementByClassNameAndIndex获取元素超时, className = " + className + ", index = " + index.ToString());
            }
            else
            {
                return null;
            }
        }
        #endregion 

        #region 根据类型、序号获取元素
        public ReadOnlyCollection<AndroidElement> GetElementsByClassName(string className, bool errorNone)
        {
            DateTime startTime = DateTime.Now;

            while ((DateTime.Now - startTime).TotalMilliseconds < OperateTimeout)
            {
                Thread.Sleep(StatusCheckInterval); 
                ReadOnlyCollection<AndroidElement> elements = Driver.FindElementsByClassName(className);
                if (elements != null && elements.Count > 0)
                {
                    return elements;
                }
            }
            if (errorNone)
            {
                throw new Exception("GetElementByClassNameAndIndex获取元素超时, className = " + className );
            }
            else
            {
                return null;
            }
        }
        #endregion

        #region 根据类型、序号获取元素
        public AndroidElement GetElementByXPath(string xPath, bool errorNone)
        {
            DateTime startTime = DateTime.Now;

            while ((DateTime.Now - startTime).TotalMilliseconds < OperateTimeout)
            {
                Thread.Sleep(StatusCheckInterval); 
                try
                {
                    AndroidElement element = Driver.FindElementByXPath(xPath);
                    if (element != null)
                    {
                        return element;
                    }
                }
                catch (NoSuchElementException ex)
                {
                    //没有此元素
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
            if (errorNone)
            {
                throw new Exception("GetElementByXPath获取元素超时, xPath = " + xPath);
            }
            else
            {
                return null;
            }
        }
        #endregion

        #region 根据id获取元素
        public AndroidElement GetElementByIdNoWaiting(AndroidElement parentElement, string id, bool errorNone)
        {
            try
            {
                return (AndroidElement)parentElement.FindElementById(id);
            }
            catch (Exception ex)
            {
                if (errorNone)
                {
                    throw ex;
                }
                else
                {
                    return null;
                }
            }
        }
        #endregion

        #region 根据id获取元素
        public AndroidElement GetElementByIdNoWaiting( string id, bool errorNone)
        {
            try
            {
                return (AndroidElement)Driver.FindElementById(id);
            }
            catch (Exception ex)
            {
                if (errorNone)
                {
                    throw ex;
                }
                else
                {
                    return null;
                }
            }
        }
        #endregion

        #region 根据id获取元素
        public  ReadOnlyCollection<AppiumWebElement> GetElementsByIdNoWaiting(AndroidElement parentElement, string id, bool errorNone)
        {
            try
            {
                return (ReadOnlyCollection<AppiumWebElement>)parentElement.FindElementsById(id);
            }
            catch (Exception ex)
            {
                if (errorNone)
                {
                    throw ex;
                }
                else
                {
                    return null;
                }
            }
        }
        #endregion

        #region 根据id获取元素
        public AndroidElement GetElementById(string id, bool errorNone)
        {
            DateTime startTime = DateTime.Now;

            while ((DateTime.Now - startTime).TotalMilliseconds < OperateTimeout)
            {
                Thread.Sleep(StatusCheckInterval);
                try
                {
                    AndroidElement element = Driver.FindElementById(id);
                    if (element != null)
                    {
                        return element;
                    }
                }
                catch (NoSuchElementException ex)
                {
                    //没有此元素
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
            if (errorNone)
            {
                throw new Exception("GetElementById获取元素超时, Id = " + id);
            }
            else
            {
                return null;
            }
        }
        #endregion


        #region 根据id获取元素
        public AndroidElement GetElementByIds(string[] ids, bool errorNone)
        {
            DateTime startTime = DateTime.Now;

            while ((DateTime.Now - startTime).TotalMilliseconds < OperateTimeout)
            {
                Thread.Sleep(StatusCheckInterval);
                foreach (string id in ids)
                {
                    try
                    {
                        AndroidElement element = Driver.FindElementById(id);
                        if (element != null)
                        {
                            return element;
                        }
                    }
                    catch (NoSuchElementException ex)
                    {
                        //没有此元素
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
            }
            if (errorNone)
            {
                throw new Exception("GetElementByIds获取元素超时, Ids = " + CommonUtil.StringArrayToString(ids, ","));
            }
            else
            {
                return null;
            }
        }
        #endregion

        #region 根据xpath获取元素
        public ReadOnlyCollection<AndroidElement> GetElementsByXPath(string xPath, bool errorNone)
        {
            DateTime startTime = DateTime.Now;

            while ((DateTime.Now - startTime).TotalMilliseconds < OperateTimeout)
            {
                Thread.Sleep(StatusCheckInterval); 
                ReadOnlyCollection<AndroidElement> elements = Driver.FindElementsByXPath(xPath);
                if (elements != null && elements.Count > 0)
                {
                    return elements;
                }
            }
            if (errorNone)
            {
                throw new Exception("GetElementsByXPath获取元素超时, xPath = " + xPath);
            }
            else
            {
                return null;
            }
        }
        #endregion

        #region 根据id获取元素
        public ReadOnlyCollection<AndroidElement> GetElementsById(string id, bool errorNone)
        {
            DateTime startTime = DateTime.Now;

            while ((DateTime.Now - startTime).TotalMilliseconds < OperateTimeout)
            {
                Thread.Sleep(StatusCheckInterval); 
                ReadOnlyCollection<AndroidElement> elements = Driver.FindElementsById(id);
                if (elements != null && elements.Count > 0)
                {
                    return elements;
                }
            }
            if (errorNone)
            {
                throw new Exception("GetElementsByXId获取元素超时, Id = " + id);
            }
            else
            {
                return null;
            }
        }
        #endregion

        #region 获取元素
        public XmlNodeList GetXmlElementsByXPath(string xPath, bool errorNone)
        {
            DateTime startTime = DateTime.Now;

            while ((DateTime.Now - startTime).TotalMilliseconds < OperateTimeout)
            {
                Thread.Sleep(StatusCheckInterval); 
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(Driver.PageSource);
                XmlNodeList nodes = xmlDoc.SelectNodes(xPath);
                if (nodes != null && nodes.Count > 0)
                {
                    return nodes;
                }
            }
            if (errorNone)
            {
                throw new Exception("GetXmlElementsByXPath获取元素超时, xPath = " + xPath);
            }
            else
            {
                return null;
            }
        }
        #endregion

        #region 获取元素
        public List<XmlNode> GetXmlElementsByXPath(string[] xPaths, bool errorNone)
        {
            DateTime startTime = DateTime.Now;

            while ((DateTime.Now - startTime).TotalMilliseconds < OperateTimeout)
            {
                Thread.Sleep(StatusCheckInterval); 
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(Driver.PageSource);
                List<XmlNode> allNodes = new List<XmlNode>();
                foreach (string xPath in xPaths)
                {
                    XmlNodeList nodes = xmlDoc.SelectNodes(xPath);
                    if (nodes != null)
                    {
                        foreach (XmlNode node in nodes)
                        {
                            allNodes.Add(node);
                        }
                    }
                }
                if (allNodes != null && allNodes.Count > 0)
                {
                    return allNodes;
                }
            }
            if (errorNone)
            {
                throw new Exception("GetXmlElementsByXPath获取元素超时, xPaths = " + CommonUtil.StringArrayToString(xPaths, ","));
            }
            else
            {
                return null;
            }
        }
        #endregion

        #region 获取元素
        public XmlNode GetXmlElementByXPath(string xPath, bool errorNone)
        {
            DateTime startTime = DateTime.Now;

            while ((DateTime.Now - startTime).TotalMilliseconds < OperateTimeout)
            {
                Thread.Sleep(StatusCheckInterval); 
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(Driver.PageSource);
                XmlNode node = xmlDoc.SelectSingleNode(xPath);
                if (node != null)
                {
                    return node;
                }
            }
            if (errorNone)
            {
                throw new Exception("GetXmlElementByXPath获取元素超时, xPath = " + xPath);
            }
            else
            {
                return null;
            }
        }
        #endregion

        #region 获取Xml根节点
        public XmlElement GetXmlRootElement()
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(Driver.PageSource);
            return xmlDoc.DocumentElement;
        }
        #endregion

        #region 获取元素
        public AndroidElement GetElementByXPath(string xPath, string[] texts, bool errorNone)
        {
            DateTime startTime = DateTime.Now;

            while ((DateTime.Now - startTime).TotalMilliseconds < OperateTimeout)
            {
                Thread.Sleep(StatusCheckInterval); 
                ReadOnlyCollection<AndroidElement> elements = Driver.FindElementsByXPath(xPath);
                if (elements != null && elements.Count > 0)
                {
                    foreach (AndroidElement element in elements)
                    {
                        foreach (string text in texts)
                        {
                            if (text == element.Text)
                            {
                                return element;
                            }
                        }
                    }
                }
            }
            if (errorNone)
            {
                throw new Exception("GetElementsByXPath获取元素超时, xPath = " + xPath);
            }
            else
            {
                return null;
            }
        }
        #endregion

        #region Click
        public void Click(Point point)
        {
            TouchAction action = new TouchAction(Driver);
            action.Press(point.X, point.Y).Release();
            action.Perform();
        }
        #endregion

        #region 根据类型、文字获取元素
        public AndroidElement GetElementByClassNameAndText(AndroidElement parentElement, string className, string[] texts, bool fullMatch, bool andCondition, bool errorNone)
        {
            DateTime startTime = DateTime.Now;

            while ((DateTime.Now - startTime).TotalMilliseconds < OperateTimeout)
            {
                Thread.Sleep(StatusCheckInterval); 
                ReadOnlyCollection<AppiumWebElement> elements = parentElement.FindElementsByClassName(className);
                if (elements.Count > 0)
                {
                    foreach (AndroidElement element in elements)
                    {
                        if (CheckElementContainText(element, texts, andCondition))
                        {
                            return element;
                        } 
                    }
                }
            }
            if (errorNone)
            {
                string allText = CommonUtil.StringArrayToString(texts, ",");
                throw new Exception("GetElementByClassNameAndText获取元素超时, className=" + className + ", text = " + allText);
            }
            else
            {
                return null;
            }
        }
        #endregion

        #region 根据类型、文字获取元素
        public AndroidElement GetElementByClassNameAndText(string className, string[] texts, bool andCondition, bool errorNone)
        {
            DateTime startTime = DateTime.Now;

            while ((DateTime.Now - startTime).TotalMilliseconds < OperateTimeout)
            {
                Thread.Sleep(StatusCheckInterval); 
                ReadOnlyCollection<AndroidElement> elements = Driver.FindElementsByClassName(className);
                if (elements.Count > 0)
                {
                    foreach (AndroidElement element in elements)
                    {
                        if (CheckElementContainText(element, texts, andCondition))
                        {
                            return element;
                        }
                    }
                }
            }
            if (errorNone)
            {
                string allText = CommonUtil.StringArrayToString(texts, ",");
                throw new Exception("GetElementByClassNameAndText获取元素超时, className=" + className + ", text = " + allText);
            }
            else
            {
                return null;
            }
        }
        #endregion

        #region 根据类型、文字获取元素
        public AndroidElement GetElementByClassNameAndText(string className, string text, bool fullMatch, bool andCondition, bool errorNone)
        {
            DateTime startTime = DateTime.Now;
            while ((DateTime.Now - startTime).TotalMilliseconds < OperateTimeout)
            {
                Thread.Sleep(StatusCheckInterval); 
                ReadOnlyCollection<AndroidElement> elements = Driver.FindElementsByClassName(className);
                if (elements.Count > 0)
                {
                    foreach (AndroidElement element in elements)
                    {
                        bool isMatch = fullMatch ? element.Text == text : element.Text.Contains(text);
                        if(isMatch)
                        {
                            return element;
                        }
                    }
                }
            }
            if (errorNone)
            { 
                throw new Exception("GetElementByClassNameAndText获取元素超时, className=" + className + ", text = " + text);
            }
            else
            {
                return null;
            }
        }
        #endregion

        #region 判断是否存在对应的文字
        private bool CheckContainText(string sourceText, string[] checkStrings, bool andCondition)
        {
            if (sourceText != null && sourceText.Length != 0)
            {
                if (andCondition)
                {
                    foreach (string checkStr in checkStrings)
                    {
                        if (!sourceText.Contains(checkStr))
                        {
                            return false;
                        }
                    }
                    return true;
                }
                else
                {
                    foreach (string checkStr in checkStrings)
                    {
                        if (sourceText.Contains(checkStr))
                        {
                            return true;
                        }
                    }
                }
            }
            return false;
        }
        #endregion

        #region 根据类型、文字获取元素
        public ReadOnlyCollection<AndroidElement> GetElementsByClassName(AndroidElement parentElement, string className, bool errorNone)
        {
            DateTime startTime = DateTime.Now;
            while ((DateTime.Now - startTime).TotalMilliseconds < OperateTimeout)
            {
                Thread.Sleep(StatusCheckInterval); 
                ReadOnlyCollection<AppiumWebElement> allChildElements = parentElement.FindElementsByClassName(className);
                if (allChildElements.Count > 0)
                {
                    List<AndroidElement> elements = new List<AndroidElement>();
                    foreach (AppiumWebElement element in allChildElements)
                    {
                        elements.Add((AndroidElement)element);
                    }
                    return new ReadOnlyCollection<AndroidElement>(elements);
                }
            }
            if (errorNone)
            {
                throw new Exception("GetElementsByClassNameAndText获取元素超时, className=" + className);
            }
            else
            {
                return null;
            }
        }
        #endregion 
        
        #region 滑动
        public void Swipe(Point fromPoint, Point toPoint, int duration)
        {
            try
            {
                Driver.Swipe(fromPoint.X, fromPoint.Y, toPoint.X, toPoint.Y, duration);
            }
            catch (Exception ex)
            {
                //经常莫名其妙的报错
            }
        }
        #endregion

        #region 点击
        public void Tap(int fingers, int x, int y, int duration)
        { 
            Driver.Tap(fingers, x, y, duration);
        }
        #endregion

        #region 发送keycode
        public void SendKeyEvent(int keyCode)
        {
            this.Driver.KeyEvent(keyCode);
        }
        #endregion


        #region 点击退回按钮
        public void ClickBackButton()
        {
            // click back button 
            this.Driver.KeyEvent(4);
        }
        #endregion

        #region 获取子节点函数
        public delegate bool PageCustomProcess(int pageIndex);
        #endregion 

        #region 滑动获取控件，直到没有再获取到新的控件（以控件包含的文字作为唯一标识）
        public void SwipeDisplayElements(Point fromPoint,
            Point toPoint, 
            int duration,
            int waitingInterval,
            PageCustomProcess processPageFunction)
        {
            List<AndroidElement> allElements = new List<AndroidElement>();
            List<string> allElementKeys = new List<string>();
            int pageIndex = 0;
            while (1 == 1)
            {
                Thread.Sleep(waitingInterval); 
                bool goon = processPageFunction(pageIndex);
                if (goon)
                { 
                    try
                    {
                        TouchAction action = new TouchAction(Driver);
                        action.Press(fromPoint.X, fromPoint.Y).Wait(duration).MoveTo(toPoint.X, toPoint.Y).Release();
                        action.Perform();
                    }
                    catch(Exception ex)
                    {
                        //此处经常运行出错，但是实际操作并没出现什么异常，比较奇怪
                    }
                }
                else
                { 
                    break;
                }
                pageIndex++;
            } 
        }
        #endregion
        
        #region 关闭
        public void Close()
        {
            if (Driver != null)
            { 
                Driver.CloseApp(); 
                Driver.Quit();
            }
        }
        #endregion
    }
}
