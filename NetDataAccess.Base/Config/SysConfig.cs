using NetDataAccess.Base.EnumTypes;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Xml;

namespace NetDataAccess.Base.Config
{
    /// <summary>
    /// 服务配置类，要改写成配置到配置文件中
    /// </summary>
    public class SysConfig
    {
        #region 常量
        public static string ListSheetName = "List";
        public static string ListPageIndexFieldName = "listPageIndex";
        public static string DetailPageUrlFieldName = "detailPageUrl";
        public static string DetailPageNameFieldName = "detailPageName";
        public static string DetailPageCookieFieldName = "cookie";
        public static string GrabStatusFieldName = "grabStatus";
        public static string GiveUpGrabFieldName = "giveUpGrab";
        //public static int ListPageIndexFieldIndex = 0;
        public static int DetailPageUrlFieldIndex = 0;
        public static int DetailPageNameFieldIndex = 1;
        public static int DetailPageCookieFieldIndex = 1;
        public static int GrabStatusFieldIndex = 3;
        public static int GiveUpGrabFieldIndex = 4;
        public static int SystemColumnCount = 5;
        public static int ColumnTitleRowCount = 1;
        #endregion

        #region ServerIP
        private static string _ServerIP = null;
        /// <summary>
        /// ServerIP
        /// </summary>
        public static string ServerIP
        {
            get
            {
                return _ServerIP;
            }
            set
            {
                _ServerIP = value;
            }
        }
        #endregion

        #region ServerPort
        private static int _ServerPort = 8081;
        /// <summary>
        /// ServerPort
        /// </summary>
        public static int ServerPort
        {
            get
            {
                return _ServerPort;
            }
            set
            {
                _ServerPort = value;
            }
        }
        #endregion

        #region 任务监视程序轮询间隔
        private static int _TaskMonitorInterval = 10 * 1000;
        /// <summary>
        /// 任务监视程序轮询间隔
        /// </summary>
        public static int TaskMonitorInterval
        {
            get
            {
                return _TaskMonitorInterval;
            }
            set
            {
                _TaskMonitorInterval = value;
            }
        }
        #endregion

        #region 获取多少页详情页数据后保存一次
        private static int _IntervalDetailPageSave = 50;
        /// <summary>
        /// 获取多少页详情页数据后保存一次
        /// </summary>
        public static int IntervalDetailPageSave
        {
            get
            {
                return _IntervalDetailPageSave;
            }
            set
            {
                _IntervalDetailPageSave = value;
            }
        }
        #endregion

        #region 最大同时执行的任务数
        /// <summary>
        /// 最大同时执行的任务数
        /// </summary>
        private static int _MaxAliveTaskNum = 3;
        /// <summary>
        /// 最大同时执行的任务数
        /// </summary>
        public static int MaxAliveTaskNum
        {
            get
            {
                return _MaxAliveTaskNum;
            }
            set
            {
                _MaxAliveTaskNum = value;
            }
        }
        #endregion

        #region 网页访问超时时间（ms）
        /// <summary>
        /// 网页访问超时时间（ms）
        /// </summary>
        private static int _WebPageRequestTimeout = 30 * 1000;
        /// <summary>
        /// 网页访问超时时间
        /// </summary>
        public static int WebPageRequestTimeout
        {
            get
            {
                return _WebPageRequestTimeout;
            }
            set
            {
                _WebPageRequestTimeout = value;
            }
        }
        #endregion

        #region 系统默认网页编码方式
        /// <summary>
        /// 系统默认网页编码方式
        /// </summary>
        private static string _WebPageEncoding = "utf-8";
        /// <summary>
        /// 系统默认网页编码方式
        /// </summary>
        public static string WebPageEncoding
        {
            get
            {
                return _WebPageEncoding;
            }
            set
            {
                _WebPageEncoding = value;
            }
        }
        #endregion

        #region 是否显示错误详情
        private static bool _AllowShowError  = true;
        /// <summary>
        /// 是否显示错误详情
        /// </summary>
        public static bool AllowShowError
        {
            get
            {
                return _AllowShowError;
            }
            set
            {
                _AllowShowError = value;
            }
        }
        #endregion

        #region 网页加载轮询等待间隔时间（ms）
        /// <summary>
        /// 网页加载轮询等待间隔时间（ms）
        /// </summary>
        private static int _WebPageRequestInterval = 2000;
        /// <summary>
        /// 网页加载轮询等待间隔时间（ms）
        /// </summary>
        public static int WebPageRequestInterval
        {
            get
            {
                return _WebPageRequestInterval;
            }
            set
            {
                _WebPageRequestInterval = value;
            }
        }
        #endregion

        #region xpath路径分隔符
        /// <summary>
        /// xpath路径分隔符
        /// </summary>
        private static string _XPathSplit = "#split#";
        /// <summary>
        /// 网页访问超时时间
        /// </summary>
        public static string XPathSplit
        {
            get
            {
                return _XPathSplit;
            }
            set
            {
                _XPathSplit = value;
            }
        }
        #endregion

        #region 显示的日志最低级别
        private static LogLevelType _ShowLogLevelType = LogLevelType.Normal;
        /// <summary>
        /// 显示的日志最低级别
        /// </summary>
        public static LogLevelType ShowLogLevelType
        {
            get
            {
                return _ShowLogLevelType;
            }
            set
            {
                _ShowLogLevelType = value;
            }
        }
        #endregion

        #region 状态刷新间隔
        private static int _IntervalShowStatus = 30000;
        /// <summary>
        /// 状态刷新间隔
        /// </summary>
        public static int IntervalShowStatus
        {
            get
            {
                return _IntervalShowStatus;
            }
            set
            {
                _IntervalShowStatus = value;
            }
        }
        #endregion

        #region 日志刷新间隔，如果联系不刷新超过次数，那么强制刷新显示日志
        private static int _IntervalShowLog = 1000;
        /// <summary>
        /// 日志刷新间隔，如果联系不刷新超过次数，那么强制刷新显示日志
        /// </summary>
        public static int IntervalShowLog
        {
            get
            {
                return _IntervalShowLog;
            }
            set
            {
                _IntervalShowLog = value;
            }
        }
        #endregion

        #region 显示最小时间间隔的日志，两次日志直接小于此时间（ms），那么不显示
        private static int _ShowLogMinTime = 500;
        /// <summary>
        /// 显示最小时间间隔的日志，两次日志直接小于此时间（ms），那么不显示
        /// </summary>
        public static int ShowLogMinTime
        {
            get
            {
                return _ShowLogMinTime;
            }
            set
            {
                _ShowLogMinTime = value;
            }
        }
        #endregion

        #region 代理出现几次错误后被弃用
        private static int _ProxyAbandonErrorTime = 3;
        /// <summary>
        /// 代理出现几次错误后被弃用
        /// </summary>
        public static int ProxyAbandonErrorTime
        {
            get
            {
                return _ProxyAbandonErrorTime;
            }
            set
            {
                _ProxyAbandonErrorTime = value;
            }
        }
        #endregion

        #region 超过NoneGotTimeout没有获取到任何页面时，系统报错，停止抓取
        private static int _NoneGotTimeout = 1000 * 60 * 30;
        /// <summary>
        /// 超过NoneGotTimeout没有获取到任何页面时，系统报错，停止抓取
        /// </summary>
        public static int NoneGotTimeout
        {
            get
            {
                return _NoneGotTimeout;
            }
            set
            {
                _NoneGotTimeout = value;
            }
        }
        #endregion

        #region 当前系统运行环境
        private static SysExecuteType _SysExecuteType = SysExecuteType.Test;
        /// <summary>
        /// 当前系统运行环境
        /// </summary>
        public static SysExecuteType SysExecuteType
        {
            get
            {
                return _SysExecuteType;
            }
            set
            {
                _SysExecuteType = value;
            }
        }
        #endregion

        #region ConfigFilePath
        /// <summary>
        /// ConfigFilePath
        /// </summary>
        public static string ConfigFilePath = Path.Combine(Path.GetDirectoryName(Application.StartupPath), "Files/Config/SysConfig.xml");
        /// <summary>
        /// ConfigFilePath
        /// </summary>
        public static string SysFileDir = Path.Combine(Path.GetDirectoryName(Application.StartupPath), "Files");
        #endregion

        #region 浏览器类型
        private static WebBrowserType _BrowserType = WebBrowserType.IE;
        public static WebBrowserType BrowserType
        {
            get
            {
                return _BrowserType;
            }
            set
            {
                _BrowserType = value;
            }
        }
        #endregion

        #region 加载配置文件
        public static void LoadSysConfig()
        {
            XmlDocument configDoc = new XmlDocument();
            configDoc.Load(ConfigFilePath);
            XmlElement rootNode = configDoc.DocumentElement;
            XmlNodeList allNodes = rootNode.SelectNodes("property");
            Dictionary<string, string> properties = new Dictionary<string, string>();
            foreach (XmlNode node in allNodes)
            {
                string name = node.Attributes["name"].Value;
                string valueStr = node.Attributes["value"].Value;
                properties.Add(name, valueStr);
            }
            SysConfig.AllowShowError = properties.ContainsKey("AllowShowError") ? bool.Parse(properties["AllowShowError"]) : SysConfig._AllowShowError;
            SysConfig.IntervalDetailPageSave = properties.ContainsKey("IntervalDetailPageSave") ? int.Parse(properties["IntervalDetailPageSave"]) : SysConfig._IntervalDetailPageSave;
            SysConfig.IntervalShowLog = properties.ContainsKey("IntervalShowLog") ? int.Parse(properties["IntervalShowLog"]) : SysConfig._IntervalShowLog;
            SysConfig.IntervalShowStatus = properties.ContainsKey("IntervalShowStatus") ? int.Parse(properties["IntervalShowStatus"]) : SysConfig._IntervalShowStatus;
            SysConfig.MaxAliveTaskNum = properties.ContainsKey("MaxAliveTaskNum") ? int.Parse(properties["MaxAliveTaskNum"]) : SysConfig._MaxAliveTaskNum;
            SysConfig.ServerIP = properties.ContainsKey("ServerIP") ? properties["ServerIP"] : SysConfig._ServerIP;
            SysConfig.ServerPort = properties.ContainsKey("ServerPort") ? int.Parse(properties["ServerPort"]) : SysConfig._ServerPort;
            SysConfig.ShowLogLevelType = properties.ContainsKey("ShowLogLevelType") ? (LogLevelType)Enum.Parse(typeof(LogLevelType), properties["ShowLogLevelType"]) : SysConfig._ShowLogLevelType;
            SysConfig.ShowLogMinTime = properties.ContainsKey("ShowLogMinTime") ? int.Parse(properties["ShowLogMinTime"]) : SysConfig._ShowLogMinTime;
            SysConfig.TaskMonitorInterval = properties.ContainsKey("TaskMonitorInterval") ? int.Parse(properties["TaskMonitorInterval"]) : SysConfig._TaskMonitorInterval;
            SysConfig.WebPageEncoding = properties.ContainsKey("WebPageEncoding") ? properties["WebPageEncoding"] : SysConfig._WebPageEncoding;
            SysConfig.WebPageRequestInterval = properties.ContainsKey("WebPageRequestInterval") ? int.Parse(properties["WebPageRequestInterval"]) : SysConfig._WebPageRequestInterval;
            SysConfig.WebPageRequestTimeout = properties.ContainsKey("WebPageRequestTimeout") ? int.Parse(properties["WebPageRequestTimeout"]) : SysConfig._WebPageRequestTimeout;
            SysConfig.XPathSplit = properties.ContainsKey("XPathSplit") ? properties["XPathSplit"] : SysConfig._XPathSplit;
            SysConfig.ProxyAbandonErrorTime = properties.ContainsKey("ProxyAbandonErrorTime") ? int.Parse(properties["ProxyAbandonErrorTime"]) : SysConfig.ProxyAbandonErrorTime;
            SysConfig.NoneGotTimeout = properties.ContainsKey("NoneGotTimeout") ? int.Parse(properties["NoneGotTimeout"]) : SysConfig.NoneGotTimeout;
            SysConfig.SysExecuteType = properties.ContainsKey("SysExecuteType") ? (SysExecuteType)Enum.Parse(typeof(SysExecuteType), properties["SysExecuteType"]) : SysConfig.SysExecuteType;
            SysConfig.BrowserType = properties.ContainsKey("BrowserType") ? (WebBrowserType)Enum.Parse(typeof(WebBrowserType), properties["BrowserType"]) : SysConfig.BrowserType;

            CefSharp.CefSettings settings = new CefSharp.CefSettings();
            settings.CachePath = "c:\\nda_config\\chromium";
            CefSharp.Cef.Initialize(settings);
        }
        #endregion
    }
}