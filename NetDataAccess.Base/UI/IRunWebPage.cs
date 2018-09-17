using NetDataAccess.Base.Definition;
using NetDataAccess.Base.EnumTypes;
using NetDataAccess.Base.DB;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;
using NetDataAccess.Base.Log;
using System.Net;
using System.Drawing;
using NetDataAccess.Base.UserAgent;
using System.Threading;
using NetDataAccess.Base.Proxy;
using NetDataAccess.Base.Browser; 

namespace NetDataAccess.Base.UI
{
    #region 运行网页抓取的基类
    /// <summary>
    /// 运行网页抓取的基类
    /// </summary>
    public interface IRunWebPage
    {
        #region 日志
        /// <summary>
        /// 日志
        /// </summary>
        RunTaskLog TaskLog { get; }
        #endregion

        #region 本次运行的任务StepId
        /// <summary>
        /// 本次运行的任务StepId
        /// </summary>
        string StepId { get; }
        #endregion

        #region 文件路径
        /// <summary>
        /// 文件路径
        /// </summary>
        string FileDir { get; }
        #endregion

        #region 本次运行待下载列表文件地址
        string ListFilePath { get; }
        #endregion

        #region 输入文件路径
        /// <summary>
        /// 输入文件路径
        /// </summary>
        string InputDir { get; }
        #endregion

        #region 输出文件路径
        /// <summary>
        /// 输出文件路径
        /// </summary>
        string OutputDir { get; }
        #endregion

        #region 运行的任务文件路径
        /// <summary>
        /// 运行的任务文件路径
        /// </summary>
        string TaskFileDir { get; }
        #endregion

        #region Project
        /// <summary>
        /// 正在执行的项目
        /// </summary>
        Proj_Main Project { get; }
        #endregion

        #region 详情页地址列表
        /// <summary>
        /// 详情页地址列表
        /// </summary>
        List<string> DetailPageUrlList { get; }
        #endregion

        #region  详情页名称列表
        /// <summary>
        /// 详情页名称列表
        /// </summary>
        List<string> DetailPageNameList { get; }
        #endregion

        #region 网页加载完成
        /// <summary>
        /// 判断网页是否加载完成，用于浏览器方式获取网页
        /// </summary>
        /// <returns></returns>
        bool CheckIsComplete(string tabName);

        /// <summary>
        /// 判断网页是否加载完成，用于浏览器方式获取网页
        /// </summary>
        /// <param name="listRow">listRow</param>
        /// <param name="dataAccessType">加载方式</param>
        /// <param name="completeChecks">判断方式</param>
        /// <param name="tabName">tabName</param>
        /// <returns></returns>
        bool CheckIsComplete(Dictionary<string, string> listRow, Proj_DataAccessType dataAccessType, Proj_CompleteCheckList completeChecks, string tabName);
        #endregion

        #region 显示网页
        /// <summary>
        /// 在浏览器中显示网页
        /// </summary>
        /// <param name="url"></param>
        IWebBrowser InvokeShowWebPage(string url, string tabName, WebBrowserType browserType, bool doFocus);

        IWebBrowser InvokeShowWebPage(string url, string tabName, WebBrowserType browserType);
        #endregion

        #region 显示进度日志
        /// <summary>
        /// 显示日志信息
        /// </summary>
        /// <param name="msg">日志信息</param>
        /// <param name="logType">日志级别</param>
        /// <param name="immediatelyShow">是否马上显示到界面</param>
        void InvokeAppendLogText(string msg, LogLevelType logType, bool immediatelyShow);
        #endregion

        #region 保存Excel文件到硬盘 放弃使用的方法
        //void SaveExcelToDisk( string fileName);
        #endregion

        #region 获取浏览器控件
        /// <summary>
        /// 获取浏览器控件
        /// </summary>
        /// <returns></returns>
        //WebBrowser GetWebBrowser();
        IWebBrowser GetWebBrowserByName(string tabName);
        #endregion

        #region 获取HTML内容
        /// <summary>
        /// 获取浏览器中显示的HTML
        /// </summary>
        /// <returns></returns>
        string InvokeGetPageHtml(string tabName);
        #endregion

        #region 获取HTML内容
        /// <summary>
        /// 获取浏览器中显示的HTML
        /// </summary>
        /// <returns></returns>
        string InvokeGetPageHtml(IWebBrowser webBrowser);
        #endregion

        #region 给控件赋值
        /// <summary>
        /// 给浏览器中的网页元素赋值
        /// </summary>
        /// <param name="id"></param>
        /// <param name="attributeName"></param>
        /// <param name="attributeValue"></param>
        void InvokeSetControlValueById(string id, string attributeName, string attributeValue, string tabName);
        #endregion

        #region 列名序号对应
        /// <summary>
        /// 输出文件列名对应列序号
        /// </summary>
        Dictionary<string, int> ColumnNameToIndex { get; }
        #endregion

        #region 获取源码页保存地址
        /// <summary>
        /// 获取源码页保存地址
        /// </summary>
        /// <param name="pageUrl"></param>
        /// <param name="dir"></param>
        /// <returns></returns>
        string GetFilePath(string pageUrl, string dir);
        #endregion

        #region 获取预处理产生的中间文件保存地址
        /// <summary>
        /// 获取预处理产生的中间文件保存地址
        /// </summary>
        /// <param name="pageUrl"></param>
        /// <param name="dir"></param>
        /// <returns></returns>
        string GetReadFilePath(string pageUrl, string dir);
        #endregion

        #region 存放预处理产生的中间文件的目录
        /// <summary>
        /// 存放预处理产生的中间文件的目录
        /// </summary>
        /// <returns></returns>
        string GetReadFileDir();
        #endregion

        #region 获取结果文件夹
        /// <summary>
        /// 获取结果文件夹
        /// </summary>
        /// <returns></returns>
        string GetExportDir();
        #endregion

        #region 获取到的Html存放目录
        /// <summary>
        /// 获取到的Html存放目录
        /// </summary>
        /// <returns></returns>
        string GetDetailSourceFileDir();
        #endregion

        #region 获取是否放弃抓取此页
        /// <summary>
        /// 获取是否放弃抓取此页
        /// </summary>
        /// <param name="listSheet"></param>
        /// <param name="pageUrl"></param>
        /// <param name="pageIndex"></param>
        /// <returns></returns>
        bool CheckGiveUpGrabPage(IListSheet listSheet, string pageUrl, int pageIndex);
        #endregion

        #region 从中间文件中读取信息
        Dictionary<string, string> ReadDetailFieldValueFromFile(string localReadFilePath);
        #endregion

        #region 从中间文件中读取信息
        List<Dictionary<string, string>> ReadDetailFieldValueListFromFile(string localReadFilePath);
        #endregion

        #region 读取下载下来的Html，加载到HtmlDocument对象中
        /// <summary>
        /// 读取下载下来的Html，加载到HtmlDocument对象中
        /// </summary>
        /// <param name="listSheet"></param>
        /// <param name="pageIndex"></param>
        /// <returns></returns>
        HtmlAgilityPack.HtmlDocument GetLocalHtmlDocument(IListSheet listSheet, int pageIndex);
        /// <summary>
        /// 读取下载下来的Html，加载到HtmlDocument对象中
        /// </summary>
        /// <param name="listSheet"></param>
        /// <param name="pageIndex"></param>
        /// <param name="encoding"></param>
        /// <returns></returns>
        HtmlAgilityPack.HtmlDocument GetLocalHtmlDocument(IListSheet listSheet, int pageIndex, Encoding encoding);
        #endregion

        #region 通过webRequest获取网页
        /// <summary>
        /// 通过webRequest获取网页
        /// </summary>
        /// <param name="pageUrl">网页地址</param>
        /// <param name="listRow">列表记录</param>
        /// <param name="needProxy">是否使用代理</param>
        /// <param name="intervalAfterLoaded">抓取后间隔时间</param>
        /// <param name="timeout">超时时间</param>
        /// <param name="encoding">编码方式</param>
        /// <param name="cookie">cookie</param>
        /// <param name="xRequestedWith">xRequestedWith</param>
        /// <param name="autoAbandonDisableProxy">自动放弃代理服务器</param>
        /// <param name="dataAccessType">dataAccessType</param>
        /// <param name="completeChecks">completeChecks</param>
        /// <returns></returns>
        string GetTextByRequest(string pageUrl, Dictionary<string, string> listRow, bool needProxy, decimal intervalAfterLoaded, int timeout, Encoding encoding, string cookie, string xRequestedWith, bool autoAbandonDisableProxy, Proj_DataAccessType dataAccessType, Proj_CompleteCheckList completeChecks, int intervalProxyRequest);
        #endregion

        #region 通过webRequest获取文件
        /// <summary>
        /// 通过webRequest获取文件
        /// </summary>
        /// <param name="pageUrl">文件地址</param>
        /// <param name="listRow">列表记录</param>
        /// <param name="needProxy">是否使用代理</param>
        /// <param name="intervalAfterLoaded">抓取后间隔时间</param>
        /// <param name="timeout">超时时间</param>
        /// <param name="autoAbandonDisableProxy">自动放弃代理服务器</param>
        /// <returns></returns>
        byte[] GetFileByRequest(string pageUrl, Dictionary<string, string> listRow, bool needProxy, decimal intervalAfterLoaded, int timeout, bool autoAbandonDisableProxy, int intervalProxyRequest);
        #endregion

        #region 通过WebBrowser获取网页
        /// <summary>
        /// 通过WebBrowser获取网页
        /// </summary>
        /// <param name="pageUrl">网页地址</param>
        /// <param name="listRow">列表记录</param>
        /// <param name="intervalAfterLoaded">抓取后间隔时间</param>
        /// <param name="timeout">超时时间</param>
        /// <param name="completeChecks">判断网页加载完毕</param> 
        /// <param name="tabName">tabName</param>
        /// <returns></returns>
        string GetDetailHtmlByWebBrowser(string pageUrl, Dictionary<string, string> listRow, decimal intervalAfterLoaded, int timeout, Proj_CompleteCheckList completeChecks, string tabName, WebBrowserType browserType);
        #endregion

        #region 保存文件
        /// <summary>
        /// 保存文件
        /// </summary>
        /// <param name="fileText"></param>
        /// <param name="localPagePath"></param>
        /// <param name="encoding"></param>
        void SaveFile(string fileText, string localPagePath, Encoding encoding);

        /// <summary>
        /// 保存文件
        /// </summary>
        /// <param name="data"></param>
        /// <param name="localPagePath"></param>
        void SaveFile(byte[] data, string localPagePath);
        #endregion

        #region 显示网页
        IWebBrowser ShowWebPage(string pageUrl, string tabName, int webRequestTimeout, bool goonWhenTimeout, WebBrowserType browserType);

        IWebBrowser ShowWebPage(string pageUrl, string tabName, int webRequestTimeout, bool goonWhenTimeout, WebBrowserType browserType, bool doFocus);
        #endregion

        #region 向网页中增加JavaScript代码，例如函数等，方便后续网页与爬取工具交互
        void InvokeAddScriptMethod(IWebBrowser webBrowser, string scriptMethodCode);
        #endregion

        #region 调用网页JavaScript脚本
        object InvokeDoScriptMethod(IWebBrowser webBrowser, string methodName, object[] parameters);
        #endregion

        #region 实现了轮询的方法判断网页JavaScript里某个值是否等于checkValue。用于异步调用后等待执行结果
        void WaitForInvokeScript(IWebBrowser webBrowser, string scriptCheckMethod, string checkValue, int invokeTimeout);
        #endregion

        #region 开始执行爬取操作
        void BeginGrab();
        #endregion

        #region 带爬取网址列表文件路径
        string ExcelFilePath { get; }
        #endregion

        #region 判断浏览器是否已经跳转到某个网页
        bool CheckWebBrowserUrl(IWebBrowser webBrowser, String checkUrl, bool fullMatch, int timeout);
        #endregion

        #region 浏览器的网页地址
        string InvokeGetWebBrowserPageUrl(IWebBrowser webBrowser);
        #endregion

        #region 利用成功和失败关键词判断是否打开了需要的页面
        bool CheckOpenRightPage(IWebBrowser webBrowser, string[] rightStrings, string[] wrongStrings, int timeout, bool andCondition);
        #endregion

        #region 判断浏览器当前浏览的网页，是否包含指定字符串
        /// <summary>
        /// 判断浏览器当前浏览的网页，是否包含指定字符串
        /// </summary>
        /// <param name="webBrowser"></param>
        /// <param name="checkStrings">待匹配的字符串（多个）</param>
        /// <param name="timeout">超过此间隔时间，系统认为网页加载失败，会抛出异常</param>
        /// <param name="andCondition">当为true时要求所有字符串都匹配，当为false时仅匹配一个就认为网页加载成功，默认为true</param>
        /// <returns></returns>
        bool CheckWebBrowserContainsForComplete(IWebBrowser webBrowser, string[] checkStrings, int timeout, bool andCondition);

        bool InvokeCheckWebBrowserContains(IWebBrowser webBrowser, string[] checkStrings, bool andCondition);
        #endregion


        #region 从中间文件读取信息
        List<string> TryGetInfoFromMiddleFile(string fileName, string fieldName);
        #endregion

        #region 将信息写入到中间文件
        void SaveInfoToMiddleFile(string fileName, string fieldName, List<string> values);
        #endregion

        #region 从中间文件读取信息
        List<Dictionary<string, string>> TryGetInfoFromMiddleFile(string fileName, string[] fieldNames);
        #endregion

        #region 将信息写入到中间文件
        void SaveInfoToMiddleFile(string fileName, string[] fieldNames, List<Dictionary<string, string>> valuesList);
        #endregion

        #region 强制重新爬取
        /// <summary>
        /// 强制重新爬取
        /// </summary>
        bool MustReGrab { get; set; }
        #endregion

        #region 滚动页面
        void InvokeScrollDocumentMethod(IWebBrowser webBrowser, Point toPoint);
        #endregion

        #region 滚动页面
        void InvokeWebBrowserGoBackMethod(IWebBrowser webBrowser);
        #endregion

        #region 关闭网页
        void CloseWebPage(string tabName);
        #endregion

        #region UserAgent列表
        UserAgents CurrentUserAgents { get; set; }
        #endregion

        #region 下载的原文件放置的文件夹
        string GetSourceFileDir(Proj_Detail_SingleLine detailPageInfo);
        #endregion

        #region 多线程抓取计数
        int CompleteGrabCount { get; set; }
        int SucceedGrabCount { get; set; }
        int AllNeedGrabCount { get; set; }
        List<int> NeedGrabIndexs { get; set; }
        List<Thread> AllGrabDetailThreads { get; set; }
        Nullable<int> GetNextGrabDetailPageIndex();
        List<string> DetailPageCookieList { get; }
        bool Grabing { get; }
        #endregion

        #region 刷新状态提示
        void RecordGrabDetailStatus(bool succeed, DateTime beginTime, DateTime endTime);
        void RefreshGrabCount(bool succeed);
        #endregion

        #region 放弃抓取
        bool GiveUpGrabPage(IListSheet listSheet, string pageUrl, int pageIndex, Exception ex);
        #endregion

        #region 代理服务器列表
        ProxyServers CurrentProxyServers { get; }
        #endregion

        #region 关闭webBrowserHtml模式下的tabPage
        void ShowTabPage(string tabName);
        #endregion
    }
    #endregion 
}
