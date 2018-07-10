using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using NetDataAccess.Base.Common;
using HtmlAgilityPack;
using NetDataAccess.Base.Definition;
using System.Xml;
using System.Threading;
using NetDataAccess.Base.EnumTypes;
using NetDataAccess.Base.Config;
using System.IO;
using System.Web;
using System.Reflection;
using NetDataAccess.Base.UI;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Net;
using Newtonsoft.Json.Linq;
using NetDataAccess.Base.Proxy;
using NetDataAccess.Base.Web;
using System.Collections;
using ICSharpCode.SharpZipLib.Zip;
using NetDataAccess.Export;
using NetDataAccess.DB;
using NetDataAccess.Base.Reader;
using NetDataAccess.Base.DB;
using Microsoft.VisualBasic.Devices;
using mshtml;
using NetDataAccess.Config;
using log4net;
using log4net.Config;
using log4net.Repository;
using NetDataAccess.Main;
using NetDataAccess.Base.DLL;
using NetDataAccess.Base.Log;
using System.Security.Cryptography;
using NetDataAccess.Base.Server;
using System.Runtime.InteropServices;
using NetDataAccess.Base.UserAgent; 

namespace NetDataAccess.Run
{
    /// <summary>
    /// 运行一个爬取任务
    /// </summary>
    [System.Runtime.InteropServices.ComVisibleAttribute(true)]
    public partial class UserControlRunGrabWebPage : UserControl, IRunWebPage
    {
        #region 外部扩展程序
        private IExternalRunWebPage _ExternalRunPage = null;
        private IExternalRunWebPage ExternalRunPage
        {
            get
            {
                return this._ExternalRunPage;
            }
            set
            {
                 this._ExternalRunPage = value;
            }
        }
        #endregion

        #region 日志
        private RunTaskLog _TaskLog = null;
        public RunTaskLog TaskLog
        {
            get
            {
                return _TaskLog;
            }
            set
            {
                _TaskLog = value;
            }
        }
        #endregion

        #region 代理服务器列表
        private ProxyServers _CurrentProxyServers = null;
        public ProxyServers CurrentProxyServers
        {
            get
            {
                return _CurrentProxyServers;
            }
            set
            {
                _CurrentProxyServers = value;
            }
        }
        #endregion

        #region UserAgent列表
        private UserAgents _CurrentUserAgents = null;
        public UserAgents CurrentUserAgents
        {
            get
            {
                return _CurrentUserAgents;
            }
            set
            {
                _CurrentUserAgents = value;
            }
        }
        #endregion

        #region 本次运行的任务StepId
        private string _StepId = "_";
        public string StepId
        {
            get
            {
                return _StepId;
            }
        }
        #endregion

        #region 本次运行待下载列表文件地址
        private string _ListFilePath = "";
        public string ListFilePath
        {
            get
            {
                return _ListFilePath;
            }
        }
        #endregion

        #region 输出目录
        private string _OutputDir = "";
        public string OutputDir
        {
            get
            {
                return _OutputDir;
            }
        }
        #endregion

        #region 中间目录
        private string _MiddleDir = "";
        public string MiddleDir
        {
            get
            {
                return _MiddleDir;
            }
        }
        #endregion

        #region 输入目录
        private string _InputDir = "";
        public string InputDir
        {
            get
            {
                return _InputDir;
            }
        }
        #endregion

        #region 本次运行的参数
        private string _Parameters = "";
        public string Parameters
        {
            get
            {
                return _Parameters;
            }
        }
        #endregion

        #region 文件路径
        public string FileDir
        {
            get
            {
                return TaskManager.FileDir;
            }
        }
        #endregion
         
        #region 运行的任务文件路径
        public string TaskFileDir
        {
            get
            {
                return Path.Combine(TaskManager.TaskFileDir,  StepId);
            }
        }
        #endregion

        #region 构造函数
        public UserControlRunGrabWebPage(Proj_Main project)
        {
            InitializeComponent();
            Create(project, false, true, "", "", "", "", "", "");
        }
        #endregion

        #region 构造函数
        public UserControlRunGrabWebPage(Proj_Main project, bool autoRun, bool popPrompt, string listFilePath, string inputDir, string middleDir, string outputDir, string parameter, string stepId)
        {
            InitializeComponent();
            Create(project, autoRun, popPrompt, listFilePath, inputDir, middleDir, outputDir, parameter, stepId);
        }
        #endregion

        #region 创建
        private void Create(Proj_Main project, bool autoRun, bool popPrompt, string listFilePath, string inputDir, string middleDir, string outputDir, string parameter, string stepId)
        {
            //。。。。。。。。。。//动态制定log文件设置有问题 20151221  
            this._StepId = stepId;
            this._ListFilePath = listFilePath;
            this._OutputDir = outputDir;
            this._MiddleDir = middleDir;
            this._InputDir = inputDir;
            this._Parameters = parameter;
            this._AutoRun = autoRun;
            this._PopPrompt = popPrompt;

            this.Project = project;
            this.Load += new EventHandler(FormWebPage_Load); 
        }
        #endregion

        private object tabLocker = new object();

        #region 创建WebBrowser对象
        //private TabPage _TabPageWebBrowser = null;

        private TabPage CreateWebBrowserTabPage(string tabName)
        {
            this.tabControlMain.TabPages.Add(tabName, tabName);
            TabPage tabPageWebBrowser = this.tabControlMain.TabPages[tabName];
            tabPageWebBrowser.Padding = new Padding(3);
            return tabPageWebBrowser;
        }
        public WebBrowser GetWebBrowserByName(string tabName)
        {
            if (this.tabControlMain.TabPages.ContainsKey(tabName))
            {
                TabPage tabPageWebBrowser = this.tabControlMain.TabPages[tabName];
                return (WebBrowser)tabPageWebBrowser.Controls[0];
            }
            else
            {
                return null;
            }
        }

        [DllImport("KERNEL32.DLL", EntryPoint = "SetProcessWorkingSetSize", SetLastError = true, CallingConvention = CallingConvention.StdCall)]
        internal static extern bool SetProcessWorkingSetSize(IntPtr pProcess, int dwMinimumWorkingSetSize, int dwMaximumWorkingSetSize);

        [DllImport("KERNEL32.DLL", EntryPoint = "GetCurrentProcess", SetLastError = true, CallingConvention = CallingConvention.StdCall)]
        internal static extern IntPtr GetCurrentProcess();

        private void RemoveOldTabPageAndBrowser(string tabName)
        {
            if (this.tabControlMain.TabPages.ContainsKey(tabName))
            {
                TabPage oldTabPage = this.tabControlMain.TabPages[tabName];
                WebBrowser oldWebBrowser = (WebBrowser)oldTabPage.Controls[0];
                if (oldWebBrowser != null)
                {
                    oldWebBrowser.Parent.Controls.Remove(oldWebBrowser);
                    oldWebBrowser.Navigating -= new WebBrowserNavigatingEventHandler(webBrowserMain_Navigating);
                    oldWebBrowser.DocumentCompleted -= new WebBrowserDocumentCompletedEventHandler(webBrowserMain_DocumentCompleted);
                    oldWebBrowser.Dispose();
                    oldWebBrowser = null; 

                    IntPtr pHandle = GetCurrentProcess();
                    SetProcessWorkingSetSize(pHandle, -1, -1);
                }

                this.tabControlMain.TabPages.RemoveByKey(tabName);

                GC.Collect();
            }
        }

        private NdaWebBrowser CreateWebBrowser(string tabName)
        {
            TabPage tabPageWebBrowser = null;
            lock (tabLocker)
            {
                this.RemoveOldTabPageAndBrowser(tabName);
                tabPageWebBrowser = CreateWebBrowserTabPage(tabName);
            }
            if (SysConfig.SysExecuteType == SysExecuteType.Test)
            {
               // this.tabControlMain.SelectedTab = tabPageWebBrowser;
            }
            NdaWebBrowser webBrowser = new NdaWebBrowser();
            webBrowser.TabName = tabName;
            webBrowser.Dock = DockStyle.Fill;
            webBrowser.ScriptErrorsSuppressed = true;
            webBrowser.AllowWebBrowserDrop = false;
            tabPageWebBrowser.Controls.Add(webBrowser);
            webBrowser.Navigating += new WebBrowserNavigatingEventHandler(webBrowserMain_Navigating);
            webBrowser.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(webBrowserMain_DocumentCompleted);
            return webBrowser;
        }
        #endregion

        #region 自启动任务
        private bool _AutoRun = false;
        public bool AutoRun
        {
            get
            {
                return _AutoRun;
            }
        }
        #endregion

        #region 是否弹出提示
        private bool _PopPrompt = true;
        public bool PopPrompt
        {
            get
            {
                return _PopPrompt;
            }
        }
        #endregion

        #region 强制重新爬取
        private bool _MustReGrab = false;
        /// <summary>
        /// 强制重新爬取
        /// </summary>
        public bool MustReGrab
        {
            get
            {
                return _MustReGrab;
            }
            set
            {
                _MustReGrab = value;
            }
        }
        #endregion

        #region 所有抓取detail的线程
        private List<Thread> _AllGrabDetailThreads = null;
        /// <summary>
        /// 所有抓取detail的线程
        /// </summary>
        public List<Thread> AllGrabDetailThreads
        {
            get
            {
                return _AllGrabDetailThreads; 
            }
            set
            {
                this._AllGrabDetailThreads = value;
            }
        }
        #endregion

        #region 最近100次抓取耗时
        private object _RefreshStatusLocker = new object();
        private List<DateTime> _AllEndTimes = new List<DateTime>();
        private List<bool> _AllSucceeds = new List<bool>();
        //private Dictionary<DateTime, DateTime> _AllSpendTimes = new Dictionary<DateTime, DateTime>();
        private void AddSpendTimeAndSucceed(DateTime startTime, DateTime endTime, bool succeed)
        {
            if (succeed)
            {
                if (_AllEndTimes.Count >= SysConfig.IntervalShowStatus)
                {
                    DateTime deleteEndTime = _AllEndTimes[0];
                    _AllEndTimes.RemoveAt(0);
                    //_AllSpendTimes.Remove(deleteEndTime);
                }
                this._AllEndTimes.Add(endTime);
                //this._AllSpendTimes.Add(endTime, startTime);
            }
            if (_AllSucceeds.Count >= SysConfig.IntervalShowStatus)
            {
                _AllSucceeds.RemoveAt(0);
            }
            this._AllSucceeds.Add(succeed);
        }

        private decimal GetSucceedPercentage()
        {
            if (this._AllSucceeds.Count == 0)
            {
                return 0;
            }
            else
            {
                int succeedCount = 0; 
                lock (_RefreshStatusLocker)
                {
                    foreach (bool succeed in this._AllSucceeds)
                    {
                        if (succeed)
                        {
                            succeedCount++;
                        }
                    }
                }
                return (decimal)(succeedCount * 100) / (decimal)this._AllSucceeds.Count;
            }
        }
        private decimal CalcAverageMinuteSucceed()
        {
            lock (_RefreshStatusLocker)
            {
                DateTime firstEndTime = DateTime.Now.AddDays(1);
                foreach (DateTime endTime in this._AllEndTimes)
                {
                    if (firstEndTime > endTime)
                    {
                        firstEndTime = endTime;
                    }
                }

                decimal m = (decimal)(DateTime.Now - firstEndTime).TotalMinutes;

                decimal avgMCount = (this._AllEndTimes.Count / m);
                return avgMCount;
            }
        }

        private decimal CalcRemainTime(decimal avgMCount, int remainCount)
        {
            return avgMCount == 0 ? decimal.MaxValue : remainCount / avgMCount;
        }

        public void RecordGrabDetailStatus(bool succeed, DateTime beginTime, DateTime endTime)
        {
            RecordStatus(succeed, "详情页抓取进度: ", beginTime, endTime);
        }

        public void RecordReadDetailStatus(bool succeed, DateTime beginTime, DateTime endTime)
        {
            RecordStatus(succeed, "详情页预处理进度: ", beginTime, endTime);
        }

        public int GetRemainGrabCount()
        {
            int remainCount = _AllNeedGrabCount - _SucceedGrabCount;
            return remainCount;
        }

        public void ShowProcessStatus()
        {
            if ((DateTime.Now - _LastShowStatusTime).TotalMilliseconds > SysConfig.IntervalShowStatus)
            {
                _LastShowStatusTime = DateTime.Now;
                decimal succeedPercentage = this.GetSucceedPercentage();
                StringBuilder sb = new StringBuilder();
                sb.Append("已共抓取:" + _SucceedGrabCount.ToString() + "个. ");
                sb.Append("最近抓取成功率:" + succeedPercentage.ToString("0.00") + "%. ");
                decimal avgMCount = this.CalcAverageMinuteSucceed();
                sb.Append("每分钟完成" + avgMCount.ToString("0.00") + "个. ");
                int remainCount = GetRemainGrabCount();
                if (remainCount > 0 && avgMCount > 0)
                {
                    long totalM = (long)CalcRemainTime(avgMCount, remainCount);
                    long remainH = 0;
                    long remainM = 0;
                    long totalH = Math.DivRem(totalM, 60, out remainM);
                    long totalD = Math.DivRem(totalH, 24, out remainH);

                    sb.Append("还需抓取:" + remainCount.ToString() + " 个. 剩余时间:" + totalD + "天" + remainH + "时" + remainM.ToString("0.00") + "分.");
                }

                if (this.CurrentProxyServers != null)
                {
                    int availableCount = this.CurrentProxyServers.GetAvailableCount();
                    sb.Append("可用的代理服务器:" + availableCount.ToString() + "个.");
                }

                int threadCount = this.AllGrabDetailThreads == null ? 0 : this.AllGrabDetailThreads.Count;
                sb.Append("可用的线程数:" + threadCount.ToString() + "个.");

                this.InvokeShowStatus(sb.ToString());
            }
        }

        private DateTime _LastShowStatusTime = DateTime.Now;
        public void RecordStatus(bool succeed, string prefix, DateTime startTime, DateTime endTime)
        {
            lock (_RefreshStatusLocker)
            {
                DateTime dt1 = DateTime.Now;
                try
                {
                    this.AddSpendTimeAndSucceed(startTime, endTime, succeed);
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                finally
                {
                    DateTime dt2 = DateTime.Now;
                    double ts = (dt2 - dt1).TotalSeconds;
                    //this.InvokeAppendLogText("RefreshStatus: " + ts.ToString("0.0000"), LogLevelType.System, true);
                }
            }
        }
        #endregion 

        #region 获取已经抓取到的列表页序号
        private int GetSucceedListPageIndex(IListSheet sheet)
        {
            return sheet.GetSucceedListPageIndex();
        }
        #endregion
         
        #region 判断是否在抓取范围
        private bool CheckInGrabRange( int rowIndex, int startPageIndex, int endPageIndex)
        {
            int userRowIndex = rowIndex + 1;
            if (startPageIndex <= 0 && endPageIndex <= 0)
            {
                return true;
            }
            else if (startPageIndex <= 0 &&endPageIndex>0)
            {
                return userRowIndex < endPageIndex;
            }
            else if (startPageIndex > 0 && endPageIndex <= 0)
            {
                return userRowIndex > startPageIndex;
            }
            else
            {
                return startPageIndex <= userRowIndex && userRowIndex <= endPageIndex;
            }
        }
        #endregion

        #region 网页加载
        private Dictionary<string, bool> _IsCompleted = new Dictionary<string,bool>();
        private Dictionary<string, bool> IsCompleted
        {
            get
            {
                return _IsCompleted;
            }
            set
            {
                _IsCompleted = value;
            }
        }
        private bool GetIsCompleted(string tabName)
        {
            return _IsCompleted.ContainsKey(tabName) ? _IsCompleted[tabName] : false;
        }

        public bool CheckIsComplete(string tabName)
        {
           return this.GetIsCompleted(tabName);
        }

        public bool CheckIsComplete(Dictionary<string, string> listRow, Proj_DataAccessType dataAccessType, Proj_CompleteCheckList completeChecks, string tabName)
        {
            string webPageHtml = this.InvokeGetPageHtml(tabName);
            if (webPageHtml == null)
            {
                return false;
            }
            else
            {
                try
                {
                    this.CheckRequestCompleteFile(webPageHtml, listRow, dataAccessType, completeChecks);
                    return true;
                }
                catch (Exception ex)
                {
                    return false;
                }
            }
            /*作废的
            switch (checkType)
            { 
                case DocumentCompleteCheckType.BrowserCompleteEvent:
                    return this.GetIsCompleted(tabName);
                case DocumentCompleteCheckType.ElementExist:
                    {
                        string webPageHtml = this.InvokeGetPageHtml(tabName);
                        if (!CommonUtil.IsNullOrBlank(webPageHtml))
                        {
                            HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                            htmlDoc.LoadHtml(webPageHtml);

                            string[] pathSections = checkElementXPath.Split(new string[] { SysConfig.XPathSplit }, StringSplitOptions.RemoveEmptyEntries);
                            foreach (string path in pathSections)
                            {
                                HtmlNode checkNode = htmlDoc.DocumentNode.SelectSingleNode(path);
                                if (checkNode != null)
                                {
                                    return true;
                                }
                            }
                            return false;
                        }
                        else
                        {
                            return false;
                        }
                    }
                case DocumentCompleteCheckType.ElementValueExist:
                    {
                        string webPageHtml = this.InvokeGetPageHtml(tabName);
                        if (!CommonUtil.IsNullOrBlank(webPageHtml))
                        {
                            HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                            htmlDoc.LoadHtml(webPageHtml);
                            string[] pathSections = checkElementXPath.Split(new string[] { SysConfig.XPathSplit }, StringSplitOptions.RemoveEmptyEntries);
                            foreach (string path in pathSections)
                            {
                                HtmlNode checkNode = htmlDoc.DocumentNode.SelectSingleNode(path);
                                if (checkNode != null && checkNode.InnerText.Trim().Length > 0)
                                {
                                    return true;
                                }
                            }
                            return false;
                        }
                        else
                        {
                            return false;
                        }
                    }
                case DocumentCompleteCheckType.ElementValueDecimal:
                    {
                        string webPageHtml = this.InvokeGetPageHtml(tabName);
                        if (!CommonUtil.IsNullOrBlank(webPageHtml))
                        {
                            HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                            htmlDoc.LoadHtml(webPageHtml);
                            string[] pathSections = checkElementXPath.Split(new string[] { SysConfig.XPathSplit }, StringSplitOptions.RemoveEmptyEntries);
                            foreach (string path in pathSections)
                            {
                                HtmlNode checkNode = htmlDoc.DocumentNode.SelectSingleNode(path);
                                if (checkNode != null && checkNode.InnerText.Trim().Length > 0)
                                {
                                    string valueStr = checkNode.InnerText.Trim();
                                    decimal value = 0;
                                    if (decimal.TryParse(valueStr, out value))
                                    {
                                        return true;
                                    } 
                                }
                            }
                            return false;
                        }
                        else
                        {
                            return false;
                        }
                    }
                default:
                    throw new Exception("暂未实现以" + checkType.ToString() + "方式不判断网页是否获取完成");
            }
            */
        }

        void webBrowserMain_Navigating(object sender, WebBrowserNavigatingEventArgs e)
        {
            NdaWebBrowser webBrowser = (NdaWebBrowser)sender;
        } 
        
        void webBrowserMain_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            NdaWebBrowser webBrowser = (NdaWebBrowser)sender;
            if (webBrowser.ReadyState == WebBrowserReadyState.Complete && !webBrowser.IsBusy)
            {
                IsCompleted[webBrowser.TabName] = true;
                this.InvokeAppendLogText("页面加载完成, tabName = " + webBrowser.TabName, LogLevelType.System, true);
                //this.webBrowserMain.Document.Window.ScrollTo(10000, 10000);
                //网页加载完成  
            }
        }
        #endregion 
         
        #region Project
        private Proj_Main _Project;
        public Proj_Main Project
        {
            get
            {
                return _Project;
            }
            set
            {
                _Project = value;
            }
        } 
        #endregion  

        #region Load时
        private void FormWebPage_Load(object sender, EventArgs e)
        {
            this.Project.Format();
            this.CreateExternalRunPage();
            this.AfterLoad();
            this.DoGrab();
        }
        #endregion

        #region 创建外部程序
        private void CreateExternalRunPage()
        {
            try
            {
                Proj_CustomProgram customProgram = (Proj_CustomProgram)this.Project.ProgramExternalRunObject;
                if (customProgram == null)
                {
                    this.ExternalRunPage = new ExternalRunWebPage();
                    this.ExternalRunPage.Init(this, null);
                }
                else
                {
                    string assemblyPath = Path.Combine(this.FileDir, "Extended\\" + customProgram.AssemblyName + ".dll");
                    Assembly assembly = Assembly.LoadFile(assemblyPath);
                    string typeName = customProgram.NamespaceName + "." + customProgram.ClassName;
                    Type type = assembly.GetType(typeName);
                    object obj = assembly.CreateInstance(typeName);
                    MethodInfo initMethod = type.GetMethod("Init", new Type[] { typeof(IRunWebPage), typeof(string) });
                    initMethod.Invoke(obj, new object[] { this, customProgram.Parameters });
                    if (obj is IExternalRunWebPage)
                    {
                        this.ExternalRunPage = (IExternalRunWebPage)obj;
                    }
                    else
                    {
                        throw new Exception("创建外部程序对象没有实现接口IExternalRunWebPage. " + this.Project.ProgramExternalRun);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("创建外部程序出错. " + this.Project.ProgramExternalRun, ex);
            }
        }
        #endregion

        #region 是否正在支持爬取
        private bool _IsGrabing = false;
        private bool IsGrabing
        {
            get 
            {
                return _IsGrabing;
            }
            set 
            {
                _IsGrabing = value;
            }
        }

        #endregion

        #region 报告放弃爬取的个数
        private void ReportGrabStatus(IListSheet listSheet)
        {
            StringBuilder s = new StringBuilder();
            int giveUpCount = 0;
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                bool giveUp = listSheet.GiveUpList[i]; 
                if (giveUp)
                {
                    giveUpCount++;
                }
            }
            s.Append("放弃爬取个数: " + giveUpCount.ToString());
            this.InvokeAppendLogText(s.ToString(), LogLevelType.System, true);
        }
        #endregion

        #region 监测进程
        private DateTime _LastGotTime = DateTime.Now;
        private bool _Grabing = false;
        public bool Grabing
        {
            get
            {
                return this._Grabing;
            }
        }
        private void GrabMonitor()
        {
            //如果长时间没有爬取到数据，那么关闭之
            _Grabing = true;
            bool goon = true;
            while (goon)
            {
                if (AutoRun)
                {
                    goon = _Grabing && (DateTime.Now - _LastGotTime).TotalMilliseconds < SysConfig.NoneGotTimeout;
                }
                else
                {
                    goon = _Grabing;
                }
                Thread.Sleep(5 * 1000);
                ShowProcessStatus();
            }

            int remainCount = GetRemainGrabCount();
            if (remainCount == 0)
            {
                InvokeShowStatus("抓取完成.");
            }
            else
            {
                string errorInfo = "长时间没有爬取到数据, 剩余" + remainCount + "个为爬取";
                InvokeShowStatus(errorInfo);
                InvokeAppendLogText(errorInfo, LogLevelType.System, true);

                if (_Grabing && AutoRun)
                {
                    //输出爬取是否成功到文件
                    TaskManager.ExportGrabResultFlag(this.StepId, this.Project.Name, false, errorInfo);

                    TaskManager.CloseTaskUI(this.StepId);
                }
            }
        }
        #endregion

        #region 启动监测进程
        private void BeginGrabMonitor()
        {
            _Grabing = true;
            Thread monitorThread = new Thread(new ThreadStart(GrabMonitor));
            monitorThread.Start();
        }
        private void EndGrabMonitor()
        {
            _Grabing = false;
        }
        #endregion

        #region 文件路径
        private string ExcelFileName
        {
            get
            {
                return CommonUtil.ProcessFileName(Project.Name, "_") + ".xlsx";
            }
        }
        private string GetPartExcelFileName(int fileIndex)
        {
            return CommonUtil.ProcessFileName(Project.Name, "_") + "_" + fileIndex + ".xlsx";
        }
        private string DBFileName
        {
            get
            {
                return CommonUtil.ProcessFileName(Project.Name, "_") + ".db";
            }
        }
        public string ExcelFilePath
        {
            get
            {
                return Path.Combine(TaskFileDir, ExcelFileName);
            }
        }
        public string GetPartExcelFilePath(int fileIndex)
        {
            return Path.Combine(TaskFileDir, this.GetPartExcelFileName(fileIndex));
        }
        #endregion

        #region 抓取
        private void Grab()
        {
            DateTime startTime = DateTime.Now;
            bool succeed = false;
            string errorInfo = "";
            IListSheet listSheet = null;
            try
            {
                //创建日志
                if (AutoRun)
                {
                    this.TaskLog = new RunTaskLog(this.StepId);                    
                } 
                string dbFilePath = Path.Combine(TaskFileDir, DBFileName);
                string newListDBFilePath = Path.Combine(FileDir, "Config\\" + "newList.db");

                this.CurrentProxyServers = new ProxyServers();
                this.CurrentUserAgents = new UserAgents();

                InvokeAppendLogText("正在创建任务...", LogLevelType.System, true); 
                
                listSheet = this.CreateListSheet(dbFilePath, newListDBFilePath); 


                this.GetDetailPageList(listSheet);

                bool hasGrabAll = true;
                if (this._AllowRunGrabDetail)
                {
                    BeginGrabMonitor();
                    hasGrabAll = false;
                    int beginGrabTime = 0;
                    while (!hasGrabAll)
                    {
                        if (beginGrabTime > 0)
                        {
                            InvokeAppendLogText("开始抓取详情页...(第" + beginGrabTime.ToString() + "次重新启动抓取任务)", LogLevelType.System, true);
                        }
                        else
                        {
                            InvokeAppendLogText("开始抓取详情页...", LogLevelType.System, true);
                        }
                        if (this.CurrentProxyServers.GetAvailableCount() == 0)
                        {
                            this.CurrentProxyServers = new ProxyServers();
                        }

                        hasGrabAll = GrabDetailPage(listSheet);
                        InvokeAppendLogText(hasGrabAll ? "详情页抓取完成!" : "未能完成详情页抓取!", LogLevelType.System, true);
                        GC.Collect();
                        beginGrabTime++;
                    }

                    //报告爬取状态
                    this.ReportGrabStatus(listSheet);

                    InvokeAppendLogText("详情页抓取完成!", LogLevelType.System, true);

                    EndGrabMonitor();
                }
                else
                {
                    InvokeAppendLogText("跳过抓取详情页.", LogLevelType.System, true);
                }

                if (hasGrabAll)
                {
                    bool hasReadAll = true;
                    if (this._AllowRunRead)
                    {
                        InvokeAppendLogText("开始读取详情页信息...", LogLevelType.System, true);
                        hasReadAll = ReadDetailPage(listSheet);
                        InvokeAppendLogText(hasReadAll ? "详情页读取完成!" : "未能完成详情页读取!", LogLevelType.System, true);
                        GC.Collect();
                    }
                    else
                    {
                        InvokeAppendLogText("跳过读取详情页.", LogLevelType.System, true);
                    }
                    if (hasReadAll)
                    {
                        bool hasExportAll = true;
                        if (this._AllowRunExport)
                        {
                            InvokeAppendLogText("开始写入到输出文件...", LogLevelType.System, true);
                            hasExportAll = ExportDetailPage(listSheet);
                            InvokeAppendLogText(hasReadAll ? "输出文件完成!" : "未能完成文件输出!", LogLevelType.System, true);
                            GC.Collect();
                        }

                        if (hasExportAll)
                        {
                            bool hasRunProcessAfterGrabAll = true;
                            if (this._AllowRunCustom)
                            {
                                InvokeAppendLogText("开始执行后期处理外部程序...", LogLevelType.System, true);
                                hasRunProcessAfterGrabAll = ProcessAfterGrabAll(listSheet);
                                InvokeAppendLogText(hasReadAll ? "后期处理外部程序执行完成!" : "未能成功执行后期处理外部程序!", LogLevelType.System, true);
                                GC.Collect();
                            }
                            if (hasRunProcessAfterGrabAll)
                            {
                                IsGrabing = false;
                                succeed = true;
                                InvokeAppendLogText("全部完成!", LogLevelType.System, true);
                            }
                            else
                            {
                                IsGrabing = false;
                            }
                        }
                        else
                        {
                            IsGrabing = false;
                        }
                    }
                    else
                    {
                        IsGrabing = false;
                    }
                }
                else
                {
                    IsGrabing = false;
                }
            }
            catch (Exception ex)
            {
                IsGrabing = false;
                errorInfo = CommonUtil.GetExceptionAllMessage(ex);
                InvokeAppendLogText("错误!" + errorInfo, LogLevelType.Error, true);

                //关闭日志
                if (AutoRun && TaskLog != null)
                {
                    TaskLog.Close();
                }
            }
            finally
            {
                if (listSheet != null)
                {
                    listSheet.Close();
                }

                DateTime endTime = DateTime.Now;

                //如果是自动运行，那么自动关闭
                if (AutoRun)
                {
                    //输出爬取是否成功到文件
                    TaskManager.ExportGrabResultFlag(this.StepId, this.Project.Name, succeed, errorInfo);

                    TaskManager.CloseTaskUI(this.StepId);
                }
            }
        }
        #endregion

        #region 显示进度
        private DateTime _LastLogTime = DateTime.Now;
        private int hiddenCount = 0;
        private delegate WebBrowser ShowWebPageInvokeDelegate(string url, string tabName);
        private delegate void CloseWebPageInvokeDelegate(string tabName);
        private delegate void SaveLogToFileDelegate(string msg, LogLevelType logLevel, bool immediatelyShow);
        private delegate void GrabInvokeDelegate(string msg);
        private delegate void ShowStatusInvokeDelegate(string msg);
        public void InvokeAppendLogText(string msg, LogLevelType logLevel, bool immediatelyShow)
        {
            if (logLevel >= SysConfig.ShowLogLevelType)
            {
                bool canShow = immediatelyShow || (DateTime.Now - _LastLogTime).TotalMilliseconds > SysConfig.ShowLogMinTime || hiddenCount > SysConfig.IntervalShowLog;
                canShow = canShow && (logLevel != LogLevelType.Error || SysConfig.AllowShowError);
                if (canShow)
                {
                    hiddenCount = 0;
                    msg = DateTime.Now.ToString("(yyyy-MM-dd HH:mm:ss) ") + logLevel.ToString() + ": " + msg;
                    this.Invoke(new GrabInvokeDelegate(this.AppendLogText), msg);

                }
                else
                {
                    hiddenCount++;
                }
                _LastLogTime = DateTime.Now;
            }
        }
        private void AppendLogText(string msg)
        {
            this.textBoxGrabLog.AppendText(msg + "\r\n");
            //记录到日志文件
            if (AutoRun)
            {
                TaskLog.AddLog(msg);
            }
        }
         
        public void InvokeShowStatus(string msg)
        {
            this.Invoke(new ShowStatusInvokeDelegate(this.ShowStatus), new object[] { msg });
        }
        private void ShowStatus(string msg)
        {
            this.textBoxStatus.Text = msg;
        }
        #endregion

        #region 抓取列表页数据
        private List<string> _DetailPageUrlList;
        public List<string> DetailPageUrlList
        {
            get
            {
                return _DetailPageUrlList;
            }
        }
        private List<string> _DetailPageCookieList;
        public List<string> DetailPageCookieList
        {
            get
            {
                return _DetailPageCookieList;
            }
        }
        private List<string> _DetailPageNameList = new List<string>();
        public List<string> DetailPageNameList
        {
            get
            {
                return _DetailPageNameList;
            }
        }

        private void GetDetailPageList(IListSheet listSheet)
        {
            listSheet.InitDetailPageInfo();
            _DetailPageUrlList = listSheet.PageUrlList;
            _DetailPageNameList = listSheet.PageNameList;
            _DetailPageCookieList = listSheet.PageCookieList;
        } 

        private bool RunExtendedProgram(Proj_CustomProgram customProgram, List<object> runParams, List<Type> runParamTypes)
        {
            try
            {
                string assemblyPath = Path.Combine(this.FileDir, "Extended\\" + customProgram.AssemblyName + ".dll");
                Assembly assembly = Assembly.LoadFile(assemblyPath);
                string typeName = customProgram.NamespaceName + "." + customProgram.ClassName;
                Type type = assembly.GetType(typeName);
                MethodInfo initMethod = type.GetMethod("Init", new Type[] { typeof(IRunWebPage) });
                object obj = assembly.CreateInstance(typeName);
                initMethod.Invoke(obj, new object[] { this });

                Type[] types ;
                if (runParamTypes == null)
                {
                    types = new Type[] { typeof(string) };
                }
                else
                {
                    runParamTypes.Insert(0, typeof(string));
                    types = runParamTypes.ToArray();
                }

                //如果为自启动任务，那么Parameter值可能不为空，那么用自启动任务的传入参数，否则用Project中定义的参数
                string parameter = CommonUtil.IsNullOrBlank(this.Parameters) ? customProgram.Parameters : this.Parameters;

                MethodInfo runMethod = type.GetMethod("Run", types);
                object[] runParamObjects;
                if (runParams == null)
                {
                    runParamObjects = new object[] { parameter };
                }
                else
                {
                    runParams.Insert(0, parameter);
                    runParamObjects = runParams.ToArray();
                }
                bool succeed = (bool)runMethod.Invoke(obj, runParamObjects);
                return succeed;
            }
            catch (Exception ex)
            {
                throw new Exception("执行外部程序出错.", ex);
            }
        }            
        #endregion

        #region 获取符合条件的节点
        private HtmlNode GetHtmlNode(HtmlNode parentNode, string allPathes) 
        {
            string[] pathes = allPathes.Split(new string[] { "#or#" }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string path in pathes)
            {
                if (path != null && path.Length > 0)
                {
                    HtmlNode node = parentNode.SelectSingleNode(path);
                    if (node != null)
                    {
                        return node;
                    }
                }
            }
            return null;
        }
        #endregion

        #region 获取路径中的参数值
        /*
        private string GetParamValue(string url, string paramterName)
        {
            if (!CommonUtil.IsNullOrBlank(paramterName))
            {

                Uri uri = new Uri(this.webBrowserMain.Url, url, false);                
                string value = HttpUtility.ParseQueryString(uri.Query).Get(paramterName);
                return CommonUtil.IsNullOrBlank(value) ? url : value;
            }
            else 
            {
                return url;
            }
        }
         */
        #endregion 

        #region 获取浏览器HTML内容
        private delegate string WebPageInvokeDelegate(string tabName);
        public string InvokeGetPageHtml(string tabName)
        {
            string html = null;
            try
            {
                html = (string)this.Invoke(new WebPageInvokeDelegate(this.GetPageHtmlContent), new object[] { tabName });
            }
            catch (Exception ex)
            { }
            return html;
        }
        private string GetPageHtmlContent(string tabName)
        {
            WebBrowser webBrowser = this.GetWebBrowserByName(tabName);
            return webBrowser.Document == null ? null : (webBrowser.Document.Body == null ? null : webBrowser.Document.Body.OuterHtml);
        }
        #endregion

        #region 抓取完成后的后期处理
        private bool ProcessAfterGrabAll(IListSheet listSheet)
        {
            return this.ExternalRunPage.AfterAllGrab(listSheet);
            /*
            Proj_CustomProgram processAfterGrabAll = (Proj_CustomProgram)this.Project.ProgramAfterGrabAllObject;
            if (processAfterGrabAll != null)
            {
                List<object> runParams = new List<object>(); 
                runParams.Add(listSheet);

                List<Type> runParamTypes = new List<Type>();
                runParamTypes.Add(typeof(IListSheet)); 

                return this.RunExtendedProgram(processAfterGrabAll, runParams, runParamTypes);
            }
            else
            {
                return true;
            }*/
        }
        #endregion

        #region 通过webrequest获取网页html

        private object _ResponseLocker = new object();
        private Dictionary<string, object> _PageToResponseString = new Dictionary<string, object>();
        private void AddResponseData(string key, object value)
        {
            lock (_ResponseLocker)
            {
                if (_PageToResponseString.ContainsKey(key))
                {
                    _PageToResponseString[key] = value;
                }
                else
                {
                    _PageToResponseString.Add(key, value);
                }
            }
        }
        private object GetResponseString(string key)
        {
            lock (_ResponseLocker)
            {
                if (_PageToResponseString.ContainsKey(key))
                {
                    return _PageToResponseString[key];
                }
                else
                {
                    return null;
                }
            }
        }
        private void RemoveResponseData(string key)
        {
            lock (_ResponseLocker)
            {
                _PageToResponseString.Remove(key);
            }
        } 

        public string GetTextByRequest(string pageUrl, Dictionary<string, string> listRow, bool needProxy, decimal intervalAfterLoaded, int timeout, Encoding encoding, string cookie, string xRequestedWith, bool autoAbandonDisableProxy, Proj_DataAccessType dataAccessType, Proj_CompleteCheckList completeChecks, int intervalProxyRequest)
        {
            NDAWebClient client = null; 
            try
            {
                DateTime dt1 = DateTime.Now;
                client = new NDAWebClient();
                client.Id = pageUrl;
                client.ResponseEncoding = encoding;
                System.Net.ServicePointManager.DefaultConnectionLimit = 512;
                client.Timeout = timeout;  
                if (needProxy)
                {
                    client.ProxyServer = this.CurrentProxyServers.BeginUse(intervalProxyRequest);
                }
                //client.Headers.Add("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; .NET CLR 1.0.3705;)");
                client.Headers.Add("user-agent", "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/46.0.2490.86 Safari/537.36");
                if (!CommonUtil.IsNullOrBlank(cookie))
                {
                    client.Headers.Add("cookie", cookie);
                    //client.Headers.Add("connection", "keep-alive");
                }
                if (!CommonUtil.IsNullOrBlank(xRequestedWith))
                {
                    client.Headers.Add("x-requested-with", xRequestedWith);
                }

                this.ExternalRunPage.WebRequestHtml_BeforeSendRequest(pageUrl, listRow, client);
                byte[] requestData = this.ExternalRunPage.GetRequestData_BeforeSendRequest(pageUrl, listRow, encoding);

                if (requestData == null)
                {
                    client.OpenReadCompleted += client_OpenReadCompleted;
                    client.OpenReadAsync(new Uri(pageUrl));
                }
                else
                {
                    client.UploadDataCompleted += client_UploadDataCompleted;
                    client.UploadDataAsync(new Uri(pageUrl), "POST", requestData);
                }

                int waitingTime = 0;
                object data = null;
                while (data == null && waitingTime < timeout)
                {
                    data = GetResponseString(client.Id);
                    if (data == null)
                    {
                        waitingTime = waitingTime + 3000;
                        Thread.Sleep(3000);
                    }
                }

                if (data != null)
                {
                    RemoveResponseData(client.Id);
                    if (data is Exception)
                    {
                        throw (Exception)data;
                    }
                    else
                    {
                        string s = null;
                        if (data is string)
                        {
                            s = (string)data;
                        }
                        if (data is byte[])
                        {
                            s = encoding.GetString((byte[])data);
                        } 

                        CheckRequestCompleteFile(s, listRow, dataAccessType, completeChecks);

                        if (needProxy)
                        {
                            this.CurrentProxyServers.Success(client.ProxyServer);
                        }

                        //再增加个等待，等待异步加载的数据
                        Thread.Sleep((int)intervalAfterLoaded);

                        DateTime dt2 = DateTime.Now;
                        double ts = (dt2 - dt1).TotalSeconds;
                        return s;
                    }
                }
                else
                {
                    throw new Exception("访问超时.");
                }
            }
            catch (NoneProxyException ex)
            {
                throw ex;
            }
            catch (Exception ex)
            {
                string errorInfo = "";
                if (needProxy)
                {
                    if (autoAbandonDisableProxy)
                    {
                        this.CurrentProxyServers.Error(client.ProxyServer);
                        if (client.ProxyServer.IsAbandon)
                        {
                            errorInfo = "放弃代理服务器:" + client.ProxyServer.IP + ":" + client.ProxyServer.Port.ToString() + ". ";
                        }
                        else
                        {
                            errorInfo = "代理服务器:" + client.ProxyServer.IP + ":" + client.ProxyServer.Port.ToString() + ". ";
                        }
                    }
                    else
                    {
                        errorInfo = "代理服务器:" + client.ProxyServer.IP + ":" + client.ProxyServer.Port.ToString() + ". ";
                    }
                }

                errorInfo = "获取网页失败.\r\n" + errorInfo + " " + pageUrl;
                throw new GrabRequestException(errorInfo, ex);
            }
            finally
            {
                if (needProxy)
                {
                    this.CurrentProxyServers.EndUse(client.ProxyServer);
                } 
            }
        }

        void client_UploadDataCompleted(object sender, UploadDataCompletedEventArgs e)
        { 
            if (_Grabing)
            {
                Stream stream = null;
                StreamReader reader = null;
                try
                {
                    NDAWebClient client = (NDAWebClient)sender;
                    if (e.Error == null)
                    { 

                        byte[] bs = e.Result;
                        AddResponseData(client.Id, bs);
                    }
                    else
                    {
                        AddResponseData(client.Id, e.Error);
                    }
                }
                catch (Exception ex)
                {
                    InvokeAppendLogText("ReadToEnd获取字符串超时. " + ex.Message, LogLevelType.System, true);
                }
                finally
                {
                    if (stream != null)
                    {
                        stream.Close();
                        stream.Dispose();
                    }
                    if (reader != null)
                    {
                        reader.Close();
                        reader.Dispose();
                    }
                }
            }
        }

        private void client_OpenReadCompleted(object sender, OpenReadCompletedEventArgs e)
        {
            if (_Grabing)
            {
                Stream stream = null;
                StreamReader reader = null;
                try
                {
                    NDAWebClient client = (NDAWebClient)sender;
                    if (e.Error == null)
                    {
                        stream = (Stream)e.Result;
                        reader = new StreamReader(stream, client.ResponseEncoding);
                        string s = reader.ReadToEnd();
                        AddResponseData(client.Id, s);
                    }
                    else
                    {
                        AddResponseData(client.Id, e.Error);
                    }
                }
                catch (Exception ex)
                {
                    InvokeAppendLogText("ReadToEnd获取字符串超时. " + ex.Message, LogLevelType.System, true);
                }
                finally
                {
                    if (stream != null)
                    {
                        stream.Close();
                        stream.Dispose();
                    }
                    if (reader != null)
                    {
                        reader.Close();
                        reader.Dispose();
                    }
                }
            }
        }
        #endregion

        #region 通过其它方式获取数据
        public void GetDataByOtherAcessType(Dictionary<string, string> listRow)
        {
            this.ExternalRunPage.GetDataByOtherAcessType(listRow);
        }
        #endregion

        #region 通过webrequest获取File
        public byte[] GetFileByRequest(string pageUrl, Dictionary<string, string> listRow, bool needProxy, decimal intervalAfterLoaded, int timeout, bool autoAbandonDisableProxy, int intervalProxyRequest)
        {
            NDAWebClient client = null;
            try
            {
                client = new NDAWebClient();
                client.Timeout = timeout;
                if (needProxy)
                {
                    ProxyServer ps = this.CurrentProxyServers.BeginUse(intervalProxyRequest);
                    client.Proxy = ps.GenerateWebProxy();
                }
                client.Headers.Add("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; .NET CLR 1.0.3705;)");
                byte[] data = client.DownloadData(pageUrl);

                if (needProxy)
                {
                    this.CurrentProxyServers.Success(client.ProxyServer);
                }

                //再增加个等待，等待异步加载的数据
                Thread.Sleep((int)intervalAfterLoaded);

                return data;
            }
            catch (NoneProxyException ex)
            {
                throw ex;
            }
            catch (Exception ex)
            {
                string errorInfo = "";
                if (needProxy)
                {
                    if (autoAbandonDisableProxy)
                    {
                        this.CurrentProxyServers.Error(client.ProxyServer);
                        if (client.ProxyServer.IsAbandon)
                        {
                            errorInfo = "放弃代理服务器:" + client.ProxyServer.IP + ":" + client.ProxyServer.Port.ToString() + ". ";
                        }
                        else
                        {
                            errorInfo = "代理服务器:" + client.ProxyServer.IP + ":" + client.ProxyServer.Port.ToString() + ". ";
                        }
                    }
                    else
                    {
                        errorInfo = "代理服务器:" + client.ProxyServer.IP + ":" + client.ProxyServer.Port.ToString() + ". ";
                    }
                }

                errorInfo = "获取文件失败.\r\n" + errorInfo;
                throw new GrabRequestException(errorInfo);
            }
            finally
            {
                if (needProxy)
                {
                    this.CurrentProxyServers.EndUse(client.ProxyServer);
                }
            }
        }
        #endregion

        #region 抓取详情页数据 
         
        private int _CompleteGrabCount = 0;
        public int CompleteGrabCount
        {
            get
            {
                return _CompleteGrabCount;
            }
            set
            {
                _CompleteGrabCount = value;
            }
        }
        private int _SucceedGrabCount = 0;
        public int SucceedGrabCount
        {
            get
            {
                return _SucceedGrabCount;
            }
            set
            {
                _SucceedGrabCount = value;
            }
        }
        private int _AllNeedGrabCount = 0;
        public int AllNeedGrabCount
        {
            get
            {
                return _AllNeedGrabCount;
            }
            set
            {
                _AllNeedGrabCount = value;
            }
        }

        private object _GetPageIndexLocker = new object();

        private List<int> _NeedGrabIndexs = null;
        public List<int> NeedGrabIndexs
        {
            get
            {
                return this._NeedGrabIndexs;
            }
            set
            {
                this._NeedGrabIndexs = value;
            }
        }
        private int InitGrabDetailPageIndexList(IListSheet listSheet, int startPageIndex, int endPageIndex, string sourceDir)
        {
            int detailPageIndex = 0;
            this._NeedGrabIndexs = new List<int>();
            this.InvokeAppendLogText("开始统计需要下载的页面.", LogLevelType.System, true);
            while (detailPageIndex < DetailPageUrlList.Count)
            {
                string pageUrl = DetailPageUrlList[detailPageIndex];
                string localPagePath = this.GetFilePath(pageUrl, sourceDir);
                if (CheckInGrabRange(detailPageIndex, startPageIndex, endPageIndex)
                    && this.CheckNeedGrab(listSheet.GetRow(detailPageIndex), localPagePath)
                    && !this.CheckGiveUpGrabPage(listSheet, pageUrl, detailPageIndex))
                {
                    this._NeedGrabIndexs.Add(detailPageIndex);
                }
                detailPageIndex++;
                if (detailPageIndex % 1000 == 0)
                {
                    double perc = (double)detailPageIndex / (double)DetailPageUrlList.Count;
                    this.InvokeAppendLogText("正在统计需要下载的页面..." + perc.ToString("#0.00%"), LogLevelType.System, true);
                }
            }
            this.InvokeAppendLogText("完成统计需要下载的页面.", LogLevelType.System, true);
            return this._NeedGrabIndexs.Count;
        }

        private bool CheckNeedGrab(Dictionary<string, string> listRow, string localPagePath)
        {
            return this.ExternalRunPage.CheckNeedGrab(listRow, localPagePath) || !this.ExistFile(localPagePath);
        }         

        public Nullable<int> GetNextGrabDetailPageIndex()
        {
            lock (_GetPageIndexLocker)
            {
                DateTime dt1 = DateTime.Now;
                try
                {
                    if (_NeedGrabIndexs.Count > 0)
                    {
                        int index = _NeedGrabIndexs[0];
                        _NeedGrabIndexs.RemoveAt(0);
                        return index;
                    }
                    else
                    {
                        return null;
                    }
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                finally
                {
                    DateTime dt2 = DateTime.Now;
                    double ts = (dt2 - dt1).TotalSeconds;
                    //this.InvokeAppendLogText("GetNextGrabDetailPageIndex: " + ts.ToString("0.0000"), LogLevelType.System, true);
                }
            }
        } 

        private void BeginGrabDetailPageInParallelThread(IListSheet listSheet, Proj_Detail_SingleLine detailPageInfo)
        {
            //int threadCount = detailPageInfo.DataAccessType == Proj_DataAccessType.WebBrowserHtml ? 1 : detailPageInfo.ThreadCount;
            int threadCount =  detailPageInfo.ThreadCount;
            this._AllGrabDetailThreads = new List<Thread>();
            for (int i = 0; i < threadCount; i++)
            {
                Thread grabThread = new Thread(new ParameterizedThreadStart(ThreadGrabDetailPageSingle));
                this.AllGrabDetailThreads.Add(grabThread);
                this.InvokeAppendLogText("线程" + grabThread.ManagedThreadId.ToString() + "开始抓取数据.", LogLevelType.System, true);
                grabThread.Start(new object[] { listSheet, detailPageInfo });
                Thread.Sleep(50);
            }
        }

        private void BeginGrabDetailPageInParallelThread(IListSheet listSheet, Proj_Detail_MultiLine detailPageInfo)
        {
            int threadCount = detailPageInfo.DataAccessType == Proj_DataAccessType.WebBrowserHtml ? 1 : detailPageInfo.ThreadCount;
            this._AllGrabDetailThreads = new List<Thread>();
            for (int i = 0; i < threadCount; i++)
            {
                Thread grabThread = new Thread(new ParameterizedThreadStart(ThreadGrabDetailPageMulti));
                this.AllGrabDetailThreads.Add(grabThread);
                this.InvokeAppendLogText("线程" + grabThread.ManagedThreadId.ToString() + "开始抓取数据.", LogLevelType.System, true);
                grabThread.Start(new object[] { listSheet, detailPageInfo });
                Thread.Sleep(50);
            }
        }

        private void ThreadGrabDetailPageSingle(object parameters)
        {
            object[] parameterArray = (object[])parameters;
            IListSheet listSheet = (IListSheet)parameterArray[0]; 
            Proj_Detail_SingleLine detailPageInfo = (Proj_Detail_SingleLine)parameterArray[1];
            string sourceDir = this.GetSourceFileDir(detailPageInfo);
            Nullable<int> nextPageIndex = this.GetNextGrabDetailPageIndex();
            while (nextPageIndex != null)
            {
                try
                { 
                    this.ThreadGrabDetailPage(listSheet, (int)nextPageIndex, detailPageInfo, sourceDir);
                }
                catch (NoneProxyException ex)
                {
                    this.InvokeAppendLogText("线程" + Thread.CurrentThread.ManagedThreadId.ToString() + ": 停止线程\r\n\r\n\r\n\r\n\r\n." + ex.Message, LogLevelType.System, true);
                    break;
                }
                catch (Exception ex)
                {
                    this.InvokeAppendLogText("线程" + Thread.CurrentThread.ManagedThreadId.ToString() + ": 出错!!!!!!!!!!!!" + ex.Message, LogLevelType.System, true);
                }
                nextPageIndex = this.GetNextGrabDetailPageIndex();
            }
            this.AllGrabDetailThreads.Remove(Thread.CurrentThread);
        }

        private void ThreadGrabDetailPageMulti(object parameters)
        {
            object[] parameterArray = (object[])parameters;
            IListSheet listSheet = (IListSheet)parameterArray[0];
            Proj_Detail_MultiLine detailPageInfo = (Proj_Detail_MultiLine)parameterArray[1];
            string sourceDir = this.GetSourceFileDir(detailPageInfo);
            Nullable<int> nextPageIndex = this.GetNextGrabDetailPageIndex();
            while (nextPageIndex != null)
            {
                try
                {
                    this.ThreadGrabDetailPage(listSheet, (int)nextPageIndex, detailPageInfo, sourceDir);
                }
                catch (NoneProxyException ex)
                {
                    this.InvokeAppendLogText("线程" + Thread.CurrentThread.ManagedThreadId.ToString() + ": 停止线程." + ex.Message, LogLevelType.System, true);
                    break;
                }
                catch (Exception ex)
                {
                    this.InvokeAppendLogText("线程" + Thread.CurrentThread.ManagedThreadId.ToString() + ": 出错!!!!!!!!!!!!" + ex.Message, LogLevelType.System, true);
                }
                nextPageIndex = this.GetNextGrabDetailPageIndex();
            }
            this.AllGrabDetailThreads.Remove(Thread.CurrentThread);
        }

        private void ThreadGrabDetailPage(IListSheet listSheet, int detailPageIndex, Proj_Detail_SingleLine detailPageInfo, string sourceDir)
        {
            DateTime dt1 = DateTime.Now;
            string pageUrl = DetailPageUrlList[detailPageIndex];
            string cookie = DetailPageCookieList[detailPageIndex];
            string localPagePath = this.GetFilePath(pageUrl, sourceDir);
            Dictionary<string, string> listRow = listSheet.GetRow(detailPageIndex);

            bool succeed = true;
            bool existLocalFile = this.ExistFile(localPagePath);

            //当抓取了一个页面/文件后
            this.ExternalRunPage.BeforeGrabOne(pageUrl, listRow, existLocalFile);

            if (!existLocalFile)
            {
                succeed = GrabDetailPage(listSheet, pageUrl, listRow, localPagePath, detailPageIndex, detailPageInfo, cookie);
            }

            //当抓取了一个页面/文件后
            this.ExternalRunPage.AfterGrabOne(pageUrl, listRow, !succeed, existLocalFile);

            this.RefreshGrabCount(succeed);

            DateTime dt2 = DateTime.Now;
            TimeSpan ts = dt2 - dt1;
            this.InvokeAppendLogText("线程" + Thread.CurrentThread.ManagedThreadId.ToString() + ": 抓取了第" + (detailPageIndex + 1).ToString() + "个页面, 用时" + ts.TotalSeconds.ToString("0.00") + "秒", LogLevelType.Normal, false);

            this.RecordGrabDetailStatus(succeed, dt1, dt2);
        }

        private void ThreadGrabDetailPage(IListSheet listSheet, int detailPageIndex, Proj_Detail_MultiLine detailPageInfo, string sourceDir)
        {
            DateTime dt1 = DateTime.Now;
            string pageUrl = DetailPageUrlList[detailPageIndex];
            string cookie = DetailPageCookieList[detailPageIndex];
            string localPagePath = this.GetFilePath(pageUrl, sourceDir);
            Dictionary<string, string> listRow = listSheet.GetRow(detailPageIndex);
            bool succeed = true;
            bool existLocalFile = this.ExistFile(localPagePath);

            //当抓取了一个页面/文件后
            this.ExternalRunPage.BeforeGrabOne(pageUrl, listRow, existLocalFile);
            if(!existLocalFile)
            {
                succeed = GrabDetailPage(listSheet, pageUrl, listRow, localPagePath, detailPageIndex, detailPageInfo, cookie);
            }

            //当抓取了一个页面/文件后
            this.ExternalRunPage.AfterGrabOne(pageUrl, listRow, !succeed, existLocalFile);

            this.RefreshGrabCount(succeed);

            DateTime dt2 = DateTime.Now;
            TimeSpan ts = dt2 - dt1;
            this.InvokeAppendLogText("线程" + Thread.CurrentThread.ManagedThreadId.ToString() + ": 抓取了第" + (detailPageIndex + 1).ToString() + "个页面, 用时" + ts.TotalSeconds.ToString("0.00") + "秒", LogLevelType.Normal, false);

            this.RecordGrabDetailStatus(succeed, dt1, dt2);
        }

        private object _GrabCounterLocker = new object();
        public void RefreshGrabCount(bool succeed)
        {
            lock (_GrabCounterLocker)
            {
                if (succeed)
                {
                    _LastGotTime = DateTime.Now;
                    _SucceedGrabCount++;
                }
                _CompleteGrabCount++;
            }
        }
        
        private bool GrabDetailPage(IListSheet listSheet)
        {
            //初始化计数器 
            _CompleteGrabCount = 0;
            _SucceedGrabCount = 0; 

            switch (this.Project.DetailGrabType)
            {
                case DetailGrabType.NoneDetailPage:
                    {
                        return true;
                    }
                case DetailGrabType.SingleLineType:
                    {
                        try
                        { 
                            Proj_Detail_SingleLine detailPageInfo = (Proj_Detail_SingleLine)this.Project.DetailGrabInfoObject;
                            string sourceDir = this.GetSourceFileDir(detailPageInfo);
                            int startPageIndex = detailPageInfo.StartPageIndex;
                            int endPageIndex = detailPageInfo.EndPageIndex;
                            _AllNeedGrabCount = this.InitGrabDetailPageIndexList(listSheet, startPageIndex, endPageIndex, sourceDir);
                            if (_AllNeedGrabCount != 0)
                            {
                                this.BeginGrabDetailPageInParallelThread(listSheet, detailPageInfo);
                                while (this._CompleteGrabCount < this._AllNeedGrabCount && this.AllGrabDetailThreads.Count > 0)
                                {
                                    Thread.Sleep(5000);
                                }
                                return this._SucceedGrabCount == this._AllNeedGrabCount;
                            }
                            else
                            {
                                return true;
                            }
                        }
                        catch (Exception ex)
                        {
                            throw ex;
                        }
                        finally
                        {
                        } 
                    }
                case DetailGrabType.MultiLineType:
                    {
                        try
                        {
                            Proj_Detail_MultiLine detailPageInfo = (Proj_Detail_MultiLine)this.Project.DetailGrabInfoObject; 
                            string sourceDir = this.GetSourceFileDir(detailPageInfo);
                            int startPageIndex = detailPageInfo.StartPageIndex;
                            int endPageIndex = detailPageInfo.EndPageIndex;
                            _AllNeedGrabCount = this.InitGrabDetailPageIndexList(listSheet, startPageIndex, endPageIndex, sourceDir);
                            if (_AllNeedGrabCount != 0)
                            {
                                this.BeginGrabDetailPageInParallelThread(listSheet, detailPageInfo);
                                while (this._CompleteGrabCount < this._AllNeedGrabCount && this.AllGrabDetailThreads.Count > 0)
                                {
                                    Thread.Sleep(5000);
                                }
                                return this._SucceedGrabCount == this._AllNeedGrabCount;
                            }
                            else
                            {
                                return true;
                            }
                        }
                        catch (Exception ex)
                        {
                            throw ex;
                        }
                        finally
                        {
                        } 
                    }
                case DetailGrabType.ProgramType:
                    { 
                        /*
                        Proj_CustomProgram customProgram = (Proj_CustomProgram)this.Project.DetailGrabInfoObject; 
                        List<object> runParams = new List<object>();
                        runParams.Add(listSheet);

                        List<Type> runParamTypes = new List<Type>();
                        runParamTypes.Add(typeof(ISheet));

                        return this.RunExtendedProgram(customProgram, runParams, runParamTypes);
                         */
                        Proj_Detail_SingleLine detailPageInfo = (Proj_Detail_SingleLine)this.Project.DetailGrabInfoObject;
                        return this.BeginGrabDetailPageInExternalProgram(listSheet, detailPageInfo);
                    } 
                default:
                    return false;
            }
        }
        #endregion

        #region 自定义ProgramType的方式，逐个抓取详情页
        private bool BeginGrabDetailPageInExternalProgram(IListSheet listSheet, Proj_Detail_SingleLine detailPageInfo)
        {
            return this.ExternalRunPage.BeginGrabDetailPageInExternalProgram(listSheet, detailPageInfo);
        }
        #endregion

        #region 创建导出文件
        private DetailExportWriter CreateExportFile(string dir, ExportType exportType, Dictionary<string, int> columnNameToIndex)
        {
            DetailExportWriter exportWriter = null;
            switch (exportType)
            {
                case ExportType.Csv:
                    {
                        string filePath = Path.Combine(dir, this.Project.Name + "_Detail.csv");
                        exportWriter = new CsvWriter(filePath, columnNameToIndex);
                    }
                    break;
                case ExportType.Excel:
                    {
                        string filePath = Path.Combine(dir, this.Project.Name + "_Detail.xlsx");
                        exportWriter = new ExcelWriter(filePath, columnNameToIndex);
                    }
                    break;
                case ExportType.None:
                    {
                        this.InvokeAppendLogText("已经设置了无需输出文件.", LogLevelType.System, true);
                        exportWriter = null;
                    }
                    break;
                case ExportType.Xml:
                    throw new Exception("尚未实现xml格式的导出功能"); 
            }
            return exportWriter;
        }
        #endregion

        #region 导出详情页信息到文件
        private bool ExportDetailPage(IListSheet listSheet)
        {
            switch (this.Project.DetailGrabType)
            {
                case DetailGrabType.NoneDetailPage:
                    {
                        return true;
                    }
                case DetailGrabType.SingleLineType:
                    { 
                        Proj_Detail_SingleLine detailPageInfo = (Proj_Detail_SingleLine)this.Project.DetailGrabInfoObject;
                        if (detailPageInfo.Fields.Count > 0)
                        {
                            int startPageIndex = detailPageInfo.StartPageIndex;
                            int endPageIndex = detailPageInfo.EndPageIndex;
                            string sourceDir = this.GetSourceFileDir(detailPageInfo);
                            string readDir = this.GetReadFileDir(detailPageInfo);

                            try
                            {
                                DetailExportWriter exportWriter = this.CreateExportFile(this.GetExportDir(detailPageInfo), detailPageInfo.ExportType, this.ColumnNameToIndex);
                                if (exportWriter != null)
                                {
                                    int detailPageIndex = 0;
                                    while (detailPageIndex < DetailPageUrlList.Count)
                                    {
                                        string pageUrl = DetailPageUrlList[detailPageIndex];
                                        string localReadFilePath = this.GetReadFilePath(pageUrl, readDir);
                                        if (this.CheckInGrabRange(detailPageIndex, startPageIndex, endPageIndex))
                                        {
                                            if (!ExportDetailPage(listSheet, exportWriter, detailPageIndex, pageUrl, localReadFilePath, detailPageInfo))
                                            {
                                                return false;
                                            }
                                        }
                                        detailPageIndex++;
                                    }
                                    exportWriter.SaveToDisk();
                                }
                                return true;
                            }
                            catch (Exception ex)
                            {
                                throw ex;
                            }
                        }
                        else
                        {
                            return true;
                        }
                    }
                case DetailGrabType.MultiLineType:
                    { 
                        Proj_Detail_MultiLine detailPageInfo = (Proj_Detail_MultiLine)this.Project.DetailGrabInfoObject;
                        if (detailPageInfo.Fields.Count > 0)
                        {
                            List<Proj_Detail_Field> allFields = GetOutputFields(detailPageInfo.Fields);
                            int startPageIndex = detailPageInfo.StartPageIndex;
                            int endPageIndex = detailPageInfo.EndPageIndex;
                            string sourceDir = this.GetSourceFileDir(detailPageInfo);
                            string readDir = this.GetReadFileDir(detailPageInfo);
                            try
                            {
                                DetailExportWriter exportWriter = this.CreateExportFile(this.GetExportDir(detailPageInfo), detailPageInfo.ExportType, this.ColumnNameToIndex);
                                if (exportWriter != null)
                                {
                                    int detailPageIndex = 0;
                                    while (detailPageIndex < DetailPageUrlList.Count)
                                    {
                                        string pageUrl = DetailPageUrlList[detailPageIndex];
                                        string localReadFilePath = this.GetReadFilePath(pageUrl, readDir);
                                        if (this.CheckInGrabRange(detailPageIndex, startPageIndex, endPageIndex))
                                        {
                                            if (!ExportDetailPage(listSheet, exportWriter, detailPageIndex, pageUrl, localReadFilePath, detailPageInfo))
                                            {
                                                return false;
                                            }
                                        }
                                        detailPageIndex++;
                                    }
                                    exportWriter.SaveToDisk();
                                }
                                return true;
                            }
                            catch (Exception ex)
                            {
                                throw ex;
                            }
                        }
                        else
                        {
                            return true;
                        }
                    }
                case DetailGrabType.ProgramType:
                    return true;
                default:
                    return false;
            }
        }
        #endregion

        #region 读取详情页信息
        private bool ReadDetailPage( IListSheet listSheet )
        {
            switch (this.Project.DetailGrabType)
            {
                case DetailGrabType.NoneDetailPage:
                    {
                        return true;
                    }
                case DetailGrabType.SingleLineType:
                    { 
                        Proj_Detail_SingleLine detailPageInfo = (Proj_Detail_SingleLine)this.Project.DetailGrabInfoObject;
                        if (detailPageInfo.Fields.Count > 0)
                        {
                            int startPageIndex = detailPageInfo.StartPageIndex;
                            int endPageIndex = detailPageInfo.EndPageIndex;
                            string sourceDir = this.GetSourceFileDir(detailPageInfo);
                            string readDir = this.GetReadFileDir(detailPageInfo);
                            bool hasError = false;
                            try
                            {
                                int detailPageIndex = 0;
                                while (detailPageIndex < DetailPageUrlList.Count)
                                {
                                    string pageUrl = DetailPageUrlList[detailPageIndex];
                                    string pageName = DetailPageNameList[detailPageIndex];
                                    string localPagePath = this.GetFilePath(pageUrl, sourceDir);
                                    string localReadFilePath = this.GetReadFilePath(pageUrl, readDir);
                                    if (CheckInGrabRange(detailPageIndex, detailPageInfo.StartPageIndex, detailPageInfo.EndPageIndex)
                                        && !this.ExistFile(localReadFilePath))
                                    {
                                        DateTime dt1 = DateTime.Now;
                                        bool succeed = ReadDetailPage(listSheet, pageUrl, pageName, localPagePath,detailPageIndex, localReadFilePath, detailPageInfo);
                                        hasError = succeed ? hasError : true;
                                        DateTime dt2 = DateTime.Now;
                                        TimeSpan ts = dt2 - dt1;
                                        this.RecordReadDetailStatus(succeed, dt1, dt2);
                                    }
                                    detailPageIndex++;
                                }
                            }
                            catch (Exception ex)
                            {
                                throw ex;
                            }
                            finally
                            {
                            }
                            return !hasError;
                        }
                        else
                        {
                            return true;
                        }
                    }
                case DetailGrabType.MultiLineType:
                    { 
                        Proj_Detail_MultiLine detailPageInfo = (Proj_Detail_MultiLine)this.Project.DetailGrabInfoObject;
                        if (detailPageInfo.Fields.Count > 0)
                        {
                            List<Proj_Detail_Field> allFields = GetOutputFields(detailPageInfo.Fields);
                            int startPageIndex = detailPageInfo.StartPageIndex;
                            int endPageIndex = detailPageInfo.EndPageIndex;
                            string sourceDir = this.GetSourceFileDir(detailPageInfo);
                            string readDir = this.GetReadFileDir(detailPageInfo);
                            bool hasError = false;
                            try
                            {
                                int detailPageIndex = 0;
                                while (detailPageIndex < DetailPageUrlList.Count)
                                {
                                    string pageUrl = DetailPageUrlList[detailPageIndex];
                                    string pageName = DetailPageNameList[detailPageIndex];
                                    string localPagePath = this.GetFilePath(pageUrl, sourceDir);
                                    string localReadFilePath = this.GetReadFilePath(pageUrl, readDir);
                                    if (CheckInGrabRange(detailPageIndex, detailPageInfo.StartPageIndex, detailPageInfo.EndPageIndex)
                                        && !this.ExistFile(localReadFilePath))
                                    {
                                        DateTime dt1 = DateTime.Now;
                                        bool succeed = ReadDetailPage(listSheet, pageUrl, pageName, localPagePath, detailPageIndex, localReadFilePath, detailPageInfo);
                                        hasError = succeed ? hasError : false;
                                        DateTime dt2 = DateTime.Now;
                                        TimeSpan ts = dt2 - dt1;
                                        //this.RefreshReadDetailStatus(succeed, dt1, dt2, detailPageInfo.StartPageIndex <= 0 ? 1 : detailPageInfo.StartPageIndex, detailPageIndex, detailPageInfo.EndPageIndex > 0 ? detailPageInfo.EndPageIndex : DetailPageUrlList.Count);
                                    }
                                    detailPageIndex++;
                                }
                            }
                            catch (Exception ex)
                            {
                                throw ex;
                            }
                            finally
                            {
                            }
                            return !hasError;
                        }
                        else
                        {
                            return true;
                        }
                    }
                case DetailGrabType.ProgramType:
                    return true;
                default:
                    return false;
            }
        }

        private bool GetNeedPartDirToSaveFile()
        {
            switch (this.Project.DetailGrabType)
            {
                case DetailGrabType.NoneDetailPage:
                    {
                        return true;
                    }
                case DetailGrabType.ProgramType:
                case DetailGrabType.SingleLineType:
                    {
                        Proj_Detail_SingleLine detailPageInfo = (Proj_Detail_SingleLine)this.Project.DetailGrabInfoObject;
                        return detailPageInfo.NeedPartDir;
                    }
                case DetailGrabType.MultiLineType:
                    {
                        Proj_Detail_MultiLine detailPageInfo = (Proj_Detail_MultiLine)this.Project.DetailGrabInfoObject;
                        return detailPageInfo.NeedPartDir;
                    }
                    /*
                case DetailGrabType.ProgramType:
                    return false;
                     */
                default:
                    return false;
            }
        }

        #region 屏蔽不友好的js，例如弹出框
        private void InvokeAvoidWebBrowserUnfriendlyJavaScript(string tabName)
        {
            WebBrowser webBrowser = this.GetWebBrowserByName(tabName);
            this.Invoke(new AvoidWebBrowserUnfriendlyJavaScriptDelegate(AvoidWebBrowserUnfriendlyJavaScript), new object[] { webBrowser });
        }
        private delegate void AvoidWebBrowserUnfriendlyJavaScriptDelegate(WebBrowser webBrowser);
        private void AvoidWebBrowserUnfriendlyJavaScript(WebBrowser webBrowser)
        { 
            HtmlElement sElement = webBrowser.Document.CreateElement("script");
            IHTMLScriptElement scriptElement = (IHTMLScriptElement)sElement.DomElement;
            scriptElement.text = "alert = function(){};confirm = function(){};";
            webBrowser.Document.Body.AppendChild(sElement); 
        }
        #endregion

        public string GetDetailHtmlByWebBrowser(string pageUrl,Dictionary<string,string> listRow, decimal intervalAfterLoaded, int timeout, Proj_CompleteCheckList completeChecks, string tabName)
        {
            InvokeShowWebPage(pageUrl, tabName);

            int waitCount = 0;
            while (!this.CheckIsComplete(listRow, Proj_DataAccessType.WebBrowserHtml, completeChecks, tabName))
            {
                if (SysConfig.WebPageRequestInterval * waitCount > timeout )
                {
                    //超时
                    throw new GrabRequestException("请求详情页超时. PageUrl = " + pageUrl);
                }
                //等待
                waitCount++;
                Thread.Sleep(SysConfig.WebPageRequestInterval);
            }

            this.ExternalRunPage.WebBrowserHtml_AfterPageLoaded(pageUrl, listRow, this.GetWebBrowserByName(tabName));

            InvokeAvoidWebBrowserUnfriendlyJavaScript(tabName);

            //再增加个等待，等待异步加载的数据
            Thread.Sleep((int)intervalAfterLoaded);
            string webPageHtml = InvokeGetPageHtml(tabName);
            return webPageHtml;
        }

        private void GetDetailHtmlInfo( string pageUrl, string pageName, string localPagePath, string localReadFilePath, Proj_Detail_SingleLine detailPageInfo, string webPageHtml)
        {
            HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
            htmlDoc.LoadHtml(webPageHtml);
            Dictionary<string, string> fieldValues = new Dictionary<string, string>();
            foreach (Proj_Detail_Field field in detailPageInfo.Fields)
            {
                if (field.Name == "*")
                {
                    fieldValues.Add(field.Name, webPageHtml);
                }
                else
                {
                    HtmlNode node = GetHtmlNode(htmlDoc.DocumentNode, field.Path);
                    string value = node == null
                        ? ""
                        : (field.NeedAllHtml ? node.OuterHtml : (CommonUtil.IsNullOrBlank(field.AttributeName) ? node.InnerText : node.GetAttributeValue(field.AttributeName, "")));
                    fieldValues.Add(field.Name, value);
                }
            }
            fieldValues.Add(SysConfig.DetailPageNameFieldName, pageName);
            fieldValues.Add(SysConfig.DetailPageUrlFieldName, pageUrl);

            this.SaveDetailFieldValueToFile( fieldValues, localReadFilePath); 

        }


        private void GetDetailJsonInfo(string pageUrl, string pageName, string localPagePath, string localReadFilePath, Proj_Detail_SingleLine detailPageInfo, string webPageText)
        {
            JObject rootJo = JObject.Parse(webPageText);  

            Dictionary<string, string> fieldValues = new Dictionary<string, string>();
            foreach (Proj_Detail_Field field in detailPageInfo.Fields)
            {
                if (field.Name == "*")
                {
                    fieldValues.Add(field.Name, rootJo.ToString());
                }
                else
                {
                    JToken fieldValueJ = rootJo[field.Name];
                    string value = fieldValueJ == null ? "" : fieldValueJ.ToString();
                    fieldValues.Add(field.Name, value);
                }
            }

            fieldValues.Add(SysConfig.DetailPageNameFieldName, pageName);
            fieldValues.Add(SysConfig.DetailPageUrlFieldName, pageUrl);
            this.SaveDetailFieldValueToFile(fieldValues, pageUrl);
        }

        private bool GrabDetailPage(IListSheet listSheet, string pageUrl,Dictionary<string,string> listRow, string localPagePath, int pageIndex, Proj_Detail_SingleLine detailPageInfo, string cookie)
        {
            string tabName = Thread.CurrentThread.ManagedThreadId.ToString();
            return this.GrabDetailPage(listSheet, pageUrl, listRow, localPagePath, pageIndex, detailPageInfo, cookie, tabName);
        }

        private bool CheckRequestCompleteFile(string webPageText, Proj_DataAccessType dataAccessType, Proj_CompleteCheck completeCheck, out string errorInfo)
        {
            errorInfo = null;
            bool succeed = false;
            DocumentCompleteCheckType checkType = completeCheck.CheckType;
            string checkValue = completeCheck.CheckValue;
            switch (checkType)
            {
                case DocumentCompleteCheckType.BrowserCompleteEvent:
                    {
                        succeed = true;
                    }
                    break;
                case DocumentCompleteCheckType.TrimEndWithText:
                    if (!webPageText.Trim().EndsWith(checkValue))
                    {
                        errorInfo = "不是以" + checkValue + "结尾，网页未获取完整. ";
                        succeed = false;
                    }
                    else
                    {
                        succeed = true;
                    }
                    break;
                case DocumentCompleteCheckType.TextExist:
                    if (!webPageText.Contains(checkValue))
                    {
                        errorInfo = "不包含" + checkValue + "，网页未获取完整. ";
                        succeed = false;
                    }
                    else
                    {
                        succeed = true;
                    }
                    break;
                default:
                    {
                        errorInfo = "暂未实现以" + checkType.ToString() + "方式不判断网页是否获取完成. ";
                        succeed = false;
                    }
                    break;
            }
            return succeed;
        }

        private void CheckRequestCompleteFile(string webPageText, Dictionary<string,string> listRow, Proj_DataAccessType dataAccessType, Proj_CompleteCheckList completeChecks)
        {
            this.ExternalRunPage.CheckRequestCompleteFile(webPageText, listRow);

            if (completeChecks != null)
            {
                if (completeChecks.AndCondition)
                {
                    foreach (Proj_CompleteCheck completeCheck in completeChecks)
                    {
                        string errorInfo = null;
                        if (!CheckRequestCompleteFile(webPageText, dataAccessType, completeCheck, out errorInfo))
                        {
                            throw new Exception(errorInfo);
                        }
                    }
                }
                else
                {
                    List<string> errorInfos = new List<string>();
                    foreach (Proj_CompleteCheck completeCheck in completeChecks)
                    {
                        string errorInfo = null;
                        if (CheckRequestCompleteFile(webPageText, dataAccessType, completeCheck, out errorInfo))
                        {
                            errorInfos.Clear();
                            break;
                        }
                        else
                        {
                            errorInfos.Add(errorInfo);
                        }
                    }
                    if (errorInfos.Count > 0)
                    {
                        throw new Exception(CommonUtil.StringArrayToString(errorInfos.ToArray(), "\r\n"));
                    }
                }
            }
        }

        private bool GrabDetailPage(IListSheet listSheet, string pageUrl, Dictionary<string,string> listRow, string localPagePath, int pageIndex, Proj_Detail_SingleLine detailPageInfo, string cookie, string tabName)
        {
            string pageName = DetailPageNameList[pageIndex];
            decimal intervalAfterLoaded = detailPageInfo.IntervalAfterLoaded;
            Encoding encoding =  Encoding.GetEncoding(detailPageInfo.Encoding);
            try
            {
                switch (detailPageInfo.DataAccessType)
                {
                    case Proj_DataAccessType.WebBrowserHtml:
                        {
                            string webPageText = GetDetailHtmlByWebBrowser(pageUrl, listRow, intervalAfterLoaded, detailPageInfo.RequestTimeout, detailPageInfo.CompleteChecks, tabName);
                            this.SaveFile(webPageText, localPagePath, encoding);
                        }
                        break;
                    case Proj_DataAccessType.WebRequestHtml:
                        {
                            string webPageText = GetTextByRequest(pageUrl, listRow, detailPageInfo.NeedProxy, intervalAfterLoaded, detailPageInfo.RequestTimeout, encoding, cookie, detailPageInfo.XRequestedWith, detailPageInfo.AutoAbandonDisableProxy, detailPageInfo.DataAccessType, detailPageInfo.CompleteChecks, detailPageInfo.IntervalProxyRequest);
                            this.SaveFile(webPageText, localPagePath, encoding);
                        }
                        break;
                    case Proj_DataAccessType.WebRequestJson:
                        {
                            string webPageText = GetTextByRequest(pageUrl, listRow, detailPageInfo.NeedProxy, intervalAfterLoaded, detailPageInfo.RequestTimeout, encoding, cookie, detailPageInfo.XRequestedWith, detailPageInfo.AutoAbandonDisableProxy, detailPageInfo.DataAccessType, detailPageInfo.CompleteChecks, detailPageInfo.IntervalProxyRequest);
                            this.SaveFile(webPageText, localPagePath, encoding);
                        }
                        break;
                    case Proj_DataAccessType.WebRequestFile:
                        {
                            byte[] data = GetFileByRequest(pageUrl, listRow, detailPageInfo.NeedProxy, intervalAfterLoaded, detailPageInfo.RequestTimeout, detailPageInfo.AutoAbandonDisableProxy, detailPageInfo.IntervalProxyRequest);
                            this.SaveFile(data, localPagePath);
                        }
                        break;
                    case Proj_DataAccessType.OtherAccessType:
                        {
                            this.GetDataByOtherAcessType(listRow);
                        }
                        break;
                }
                return true;
            }
            catch (NoneProxyException ex)
            {
                throw ex;
            }
            catch (Exception ex)
            {
                //当抓取了一个页面/文件出错后
                if (this.ExternalRunPage.AfterGrabOneCatchException(pageUrl, listRow, ex))
                {
                    string errorMsg = "线程" + Thread.CurrentThread.ManagedThreadId.ToString() + ": 放弃抓取（外部程序要求放弃）.";
                    listSheet.SetGiveUp(pageIndex, pageUrl, errorMsg);
                    this.InvokeAppendLogText(errorMsg + " PageUrl = " + pageUrl + ". " + ex.Message + (ex.InnerException == null ? "" : ex.InnerException.Message), LogLevelType.Error, true);
                    return true;
                }
                else if (!detailPageInfo.AllowAutoGiveUp || !GiveUpGrabPage(listSheet, pageUrl, pageIndex, ex))
                {
                    this.InvokeAppendLogText("线程" + Thread.CurrentThread.ManagedThreadId.ToString() + ": PageUrl = " + pageUrl + ". " + ex.Message + (ex.InnerException == null ? "" : ex.InnerException.Message), LogLevelType.Error, true);
                    return false;
                }
                else
                {
                    this.InvokeAppendLogText("线程" + Thread.CurrentThread.ManagedThreadId.ToString() + ": 放弃抓取. PageUrl = " + pageUrl + ". " + ex.Message + (ex.InnerException == null ? "" : ex.InnerException.Message), LogLevelType.Error, true);
                    return true;
                }
            }
        }

        /// <summary>
        ///  获取是否放弃抓取此页
        /// </summary>
        /// <returns></returns>
        public bool CheckGiveUpGrabPage(IListSheet listSheet, string pageUrl, int pageIndex)
        {
            bool giveUp = listSheet.GiveUpList[pageIndex];
            string url = listSheet.PageUrlList[pageIndex];
            if (url == pageUrl)
            {
                return giveUp;
            }
            else
            {
                throw new Exception("记录行定位错误. PageUrl = " + pageUrl + ". UrlCellValue = " + url);
            }
        }

        /// <summary>
        ///  判断是否放弃抓取此页
        /// </summary>
        /// <returns></returns>
        public bool GiveUpGrabPage(IListSheet listSheet, string pageUrl, int pageIndex, Exception ex)
        {
            bool giveUp = false;
            string errorMsg = "";
            if (ex is GrabRequestException)
            {
                giveUp = true;
                errorMsg = ex.Message;
            }

            if (giveUp)
            {
                listSheet.SetGiveUp(pageIndex, pageUrl, errorMsg);
                return true;
            }
            else
            {
                return false;
            }
        }


        private bool ReadDetailPage(IListSheet listSheet, string pageUrl, string pageName, string localPagePath, int detailPageIndex, string localReadFilePath, Proj_Detail_SingleLine detailPageInfo)
        {
            try
            {
                if (this.CheckGiveUpGrabPage(listSheet, pageUrl, detailPageIndex))
                {
                    return true;
                }
                else
                {
                    string webPageText = this.ReadFile(localPagePath);
                    switch (detailPageInfo.DataAccessType)
                    {
                        case Proj_DataAccessType.WebBrowserHtml:
                            {
                                this.GetDetailHtmlInfo(pageUrl, pageName, localPagePath, localReadFilePath, detailPageInfo, webPageText);
                            }
                            break;
                        case Proj_DataAccessType.WebRequestHtml:
                            {
                                this.GetDetailHtmlInfo(pageUrl, pageName, localPagePath, localReadFilePath, detailPageInfo, webPageText);
                            }
                            break;
                        case Proj_DataAccessType.WebRequestJson:
                            {
                                this.GetDetailJsonInfo(pageUrl, pageName, localPagePath, localReadFilePath, detailPageInfo, webPageText);
                            }
                            break;
                    }
                    return true;
                }
            }
            catch (Exception ex)
            {
                this.InvokeAppendLogText("读取详情页信息出错. PageUrl = " + pageUrl + " " + ex.Message, LogLevelType.Error, true);
                return false;
            }
        }

        private bool ReadDetailPage(IListSheet listSheet, string pageUrl, string pageName, string localPagePath, int detailPageIndex, string localReadFilePath, Proj_Detail_MultiLine detailPageInfo)
        {
            try
            {
                if (this.CheckGiveUpGrabPage(listSheet, pageUrl, detailPageIndex))
                {
                    return true;
                }
                else
                {
                    string webPageText = this.ReadFile(localPagePath);
                    switch (detailPageInfo.DataAccessType)
                    {
                        case Proj_DataAccessType.WebBrowserHtml:
                            {
                                this.GetDetailHtmlInfo(pageUrl, pageName, localPagePath, localReadFilePath, detailPageInfo, webPageText);
                            }
                            break;
                        case Proj_DataAccessType.WebRequestHtml:
                            {
                                this.GetDetailHtmlInfo(pageUrl, pageName, localPagePath, localReadFilePath, detailPageInfo, webPageText);
                            }
                            break;
                        case Proj_DataAccessType.WebRequestJson:
                            {
                                this.GetDetailJsonInfo(pageUrl, pageName, localPagePath, localReadFilePath, detailPageInfo, webPageText);
                            }
                            break;
                    }
                    return true;
                }
            }
            catch (Exception ex)
            {
                this.InvokeAppendLogText("读取详情页信息出错. PageUrl = " + pageUrl, LogLevelType.Error, true);
                return false;
            }
        }
        private bool ExportDetailPage(IListSheet listSheet, DetailExportWriter ew, int detailPageIndex, string pageUrl, string localReadFilePath, Proj_Detail_MultiLine detailPageInfo)
        {
            try
            {
                if (this.CheckGiveUpGrabPage(listSheet, pageUrl, detailPageIndex))
                {
                    return true;
                }
                else
                {
                    string webPageText = this.ReadFile(localReadFilePath);
                    List<Dictionary<string, string>> fieldValueList = this.ReadDetailFieldValueListFromFile(localReadFilePath);

                    foreach (Dictionary<string, string> fieldValues in fieldValueList)
                    {
                        ew.SaveDetailFieldValue(listSheet, this.ColumnNameToIndex, fieldValues, detailPageIndex, pageUrl);
                    }
                    return true;
                }
            }
            catch (Exception ex)
            {
                this.InvokeAppendLogText("写入到输出文件出错. PageUrl = " + pageUrl + " " + ex.Message, LogLevelType.Error, true);
                return false;
            }
        }
        private bool ExportDetailPage(IListSheet listSheet, DetailExportWriter ew, int detailPageIndex, string pageUrl, string localReadFilePath, Proj_Detail_SingleLine detailPageInfo)
        {
            try
            {
                if (this.CheckGiveUpGrabPage(listSheet, pageUrl, detailPageIndex))
                {
                    return true;
                }
                else
                {
                    string webPageText = this.ReadFile(localReadFilePath);
                    Dictionary<string, string> fieldValues = this.ReadDetailFieldValueFromFile(localReadFilePath);
                    ew.SaveDetailFieldValue(listSheet, this.ColumnNameToIndex, fieldValues, detailPageIndex, pageUrl);
                    return true;
                }
            }
            catch (Exception ex)
            {
                this.InvokeAppendLogText("写入到输出文件出错. PageUrl = " + pageUrl + " " + ex.Message, LogLevelType.Error, true);
                return false;
            }
        } 

        public string GetSourceFileDir(Proj_Detail_SingleLine detailPageInfo)
        {
            if (!CommonUtil.IsNullOrBlank(this.MiddleDir))
            {
                return this.MiddleDir;
            }
            else
            {
                return Path.Combine(detailPageInfo.SaveFileDirectory, "Detail");
            }
        }

        private string GetSourceFileDir(Proj_CustomProgram detailPageInfo)
        {
            if (!CommonUtil.IsNullOrBlank(this.MiddleDir))
            {
                return this.MiddleDir;
            }
            else
            {
                return Path.Combine(detailPageInfo.SaveFileDirectory, "Detail");
            }
        }

        private string GetExportDir(Proj_Detail_SingleLine detailPageInfo)
        {
            if (!CommonUtil.IsNullOrBlank(this.OutputDir))
            {
                return this.OutputDir;
            }
            else
            {
                return Path.Combine(detailPageInfo.SaveFileDirectory, "Export");
            }
        }

        private string GetExportDir(Proj_CustomProgram detailPageInfo)
        {
            if (!CommonUtil.IsNullOrBlank(this.OutputDir))
            {
                return this.OutputDir;
            }
            else
            {
                return Path.Combine(detailPageInfo.SaveFileDirectory, "Export");
            }
        }

        private string GetExportDir(Proj_Detail_MultiLine detailPageInfo)
        {
            if (!CommonUtil.IsNullOrBlank(this.OutputDir))
            {
                return this.OutputDir;
            }
            else
            {
                return Path.Combine(detailPageInfo.SaveFileDirectory, "Export");
            }
        }  

        public string GetDetailSourceFileDir()
        {
            switch (this.Project.DetailGrabType)
            {
                case DetailGrabType.MultiLineType:
                    return this.GetSourceFileDir((Proj_Detail_MultiLine)this.Project.DetailGrabInfoObject);
                case DetailGrabType.SingleLineType:
                case DetailGrabType.ProgramType:
                    return this.GetSourceFileDir((Proj_Detail_SingleLine)this.Project.DetailGrabInfoObject);
                    /*
                case DetailGrabType.ProgramType:
                    return this.GetSourceFileDir((Proj_CustomProgram)this.Project.DetailGrabInfoObject);
                     */
                default:
                    return "";
            }
        }

        public string GetExportDir()
        {
            switch (this.Project.DetailGrabType)
            {
                case DetailGrabType.MultiLineType:
                    return this.GetExportDir((Proj_Detail_MultiLine)this.Project.DetailGrabInfoObject);
                case DetailGrabType.SingleLineType:
                case DetailGrabType.ProgramType:
                    return this.GetExportDir((Proj_Detail_SingleLine)this.Project.DetailGrabInfoObject);
                    /*
                case DetailGrabType.ProgramType:
                    return this.GetExportDir((Proj_CustomProgram)this.Project.DetailGrabInfoObject);
                     */
                default:
                    return "";
            }
        }

        public string GetReadFileDir()
        {
            switch (this.Project.DetailGrabType)
            {
                case DetailGrabType.MultiLineType:
                    return this.GetReadFileDir((Proj_Detail_MultiLine)this.Project.DetailGrabInfoObject);
                case DetailGrabType.SingleLineType:
                    return this.GetReadFileDir((Proj_Detail_SingleLine)this.Project.DetailGrabInfoObject);
                default:
                    return "";
            }
        }

        private string GetReadFileDir(Proj_Detail_SingleLine detailPageInfo)
        {
            if (!CommonUtil.IsNullOrBlank(this.MiddleDir))
            {
                return Path.Combine(this.MiddleDir, "ReadDetail");
            }
            else
            {
                return Path.Combine(detailPageInfo.SaveFileDirectory, "ReadDetail");
            }
        }

        private string GetSourceFileDir(Proj_Detail_MultiLine detailPageInfo)
        {
            if (!CommonUtil.IsNullOrBlank(this.MiddleDir))
            {
                return Path.Combine(this.MiddleDir, "ReadDetail");
            }
            else
            {
                return Path.Combine(detailPageInfo.SaveFileDirectory, "Detail");
            }
        }

        private string GetReadFileDir(Proj_Detail_MultiLine detailPageInfo)
        {
            if (!CommonUtil.IsNullOrBlank(this.MiddleDir))
            {
                return Path.Combine(this.MiddleDir, "ReadDetail");
            }
            else
            {
                return Path.Combine(detailPageInfo.SaveFileDirectory, "ReadDetail");
            }
        }

        private bool ExistFile(string localPagePath)
        {
            try
            {
                bool isExist = File.Exists(localPagePath);
                /*
                string np = localPagePath + ".txt";
                if (File.Exists(np))
                {
                    File.Move(np, localPagePath);
                    isExist = true;
                }
                 */
                return isExist;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private string ReadFile(string filePath)
        {  
            TextReader tr = null;
            try
            {
                tr = new StreamReader(filePath);
                string fileText = tr.ReadToEnd();
                return fileText;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (tr != null)
                {
                    tr.Dispose();
                    tr = null;
                }
            }
        }

        public string GetReadFilePath(string pageUrl, string dir)
        {
            return GetFilePath(pageUrl + ".txt", dir);
        }

        private string GetPartDir(String fileName)
        {
            if (this.GetNeedPartDirToSaveFile())
            {
                MD5 md5 = new MD5CryptoServiceProvider();
                byte[] result = Encoding.Default.GetBytes(fileName);
                byte[] output = md5.ComputeHash(result);
                string newFileName = BitConverter.ToString(output);
                string partDirName = newFileName.Substring(0, 2);
                return partDirName;
            }
            else
            {
                return null;
            }
        }

        public string GetFilePath(string pageUrl, string dir)
        {
            try
            {
                string fileName = CommonUtil.ProcessFileName(pageUrl, "_");

                string partDirName = this.GetPartDir(fileName);

                string filePath = Path.Combine(dir, (partDirName == null ? "" : partDirName + "\\") + fileName);

                if (filePath.Length > 200)
                {
                    MD5 md5 = new MD5CryptoServiceProvider();
                    byte[] result = Encoding.Default.GetBytes(fileName);
                    byte[] output = md5.ComputeHash(result);
                    string newFileName = BitConverter.ToString(output);
                    return GetFilePath(newFileName, dir);
                }
                else
                {
                    return filePath;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void SaveFile(string fileText, string localPagePath, Encoding encoding)
        {
            if (CommonUtil.IsNullOrBlank(fileText))
            {
                throw new EmptyFileException("空文件.");
            }
            else
            {

                CommonUtil.CreateFileDirectory(localPagePath);
                StreamWriter sw = null;
                try
                {
                    sw = new StreamWriter(localPagePath, false, encoding);
                    sw.Write(fileText);
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                finally
                {
                    if (sw != null)
                    {
                        sw.Dispose();
                        sw = null;
                    }
                }
            }
        }

        public void SaveFile(byte[] data, string localPagePath)
        {
            if (data == null && data.Length>0)
            {
                throw new Exception("空文件. LocalPagePath = " + localPagePath);
            }
            else
            {
                CommonUtil.CreateFileDirectory(localPagePath);
                FileStream fs = null;
                try
                { 
                    fs = new FileStream(localPagePath, FileMode.Create, FileAccess.Write);
                    fs.Write(data, 0, data.Length);
                    fs.Flush();
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                finally
                {
                    if (fs != null)
                    {
                        fs.Close();
                        fs.Dispose();
                        fs = null;
                    }
                }
            }
        }

        private void GetDetailHtmlInfo(string pageUrl, string pageName, string localPagePath, string localReadFilePath, Proj_Detail_MultiLine detailPageInfo, string webPageHtml)
        {
            HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
            htmlDoc.LoadHtml(webPageHtml);

            string[] blockPathSections = detailPageInfo.MultiCtrlPath.Split(new string[] { SysConfig.XPathSplit }, StringSplitOptions.RemoveEmptyEntries);
            List<HtmlNode> allBlockNodes = new List<HtmlNode>();
            allBlockNodes.Add(htmlDoc.DocumentNode);
            foreach (string blockPath in blockPathSections)
            {
                if (blockPath != SysConfig.XPathSplit)
                {
                    List<HtmlNode> allNewBlockNodes = new List<HtmlNode>();
                    foreach (HtmlNode parentNode in allBlockNodes)
                    {
                        HtmlNodeCollection children = parentNode.SelectNodes(blockPath);
                        if (children != null)
                        {
                            allNewBlockNodes.AddRange(children);
                        }
                    }
                    allBlockNodes = allNewBlockNodes;
                }
            }

            List<Dictionary<string, string>> filedValuesList = new List<Dictionary<string, string>>();
            //foreach (HtmlNode relativeNode in allBlockNodes)
            for (int i = 0; i < allBlockNodes.Count; i++)
            {
                HtmlNode relativeNode = allBlockNodes[i];
                Dictionary<string, string> fieldValues = new Dictionary<string, string>();
                foreach (Proj_Detail_Field field in detailPageInfo.Fields)
                {
                    HtmlNode baseNode = field.IsAbsolute ? htmlDoc.DocumentNode : relativeNode;

                    HtmlNode node = GetHtmlNode(baseNode, field.Path);
                    string value = node == null
                        ? ""
                        : (field.NeedAllHtml ? node.OuterHtml : (CommonUtil.IsNullOrBlank(field.AttributeName) ? node.InnerText : node.GetAttributeValue(field.AttributeName, "")));
                    fieldValues.Add(field.Name, value);
                }

                fieldValues.Add(SysConfig.DetailPageNameFieldName, pageName);
                fieldValues.Add(SysConfig.DetailPageUrlFieldName, pageUrl);

                filedValuesList.Add(fieldValues);
            }

            this.SaveDetailFieldValueToFile(filedValuesList, localReadFilePath);
        }

        private void GetDetailJsonInfo(string pageUrl, string pageName, string localPagePath, string localReadFilePath, Proj_Detail_MultiLine detailPageInfo, string webPageText)
        {
            string blockPath = detailPageInfo.MultiCtrlPath;
            Object jt = null;
            if (CommonUtil.IsNullOrBlank(blockPath))
            {
                jt = JArray.Parse(webPageText);
            }
            else
            {
                JObject rootJo = JObject.Parse(webPageText);
                jt = rootJo.SelectToken(blockPath);
            }
            if (jt == null)
            {
                throw new Exception("找不到路径" + blockPath + "的值. JSON = " + webPageText);
            }
            else if (!(jt is JArray))
            {
                throw new Exception("找到了" + blockPath + "的值, 但是它不是数组. JSON = " + webPageText);
            }
            else
            {
                List<Dictionary<string, string>> filedValuesList = new List<Dictionary<string, string>>();
                JArray jas = jt as JArray;
                for (int i = 0; i < jas.Count; i++)
                {
                    Dictionary<string, string> fieldValues = new Dictionary<string, string>();
                    if (jas[i] is JValue)
                    {
                        JValue jo = jas[i] as JValue;

                        foreach (Proj_Detail_Field field in detailPageInfo.Fields)
                        {
                            if (field.Name == "*")
                            {
                                fieldValues.Add(field.Name, jo.ToString());
                            }
                            else
                            {
                                fieldValues.Add(field.Name, "");
                            }
                        }
                    }
                    else if (jas[i] is JArray)
                    {
                        JArray jo = jas[i] as JArray;

                        foreach (Proj_Detail_Field field in detailPageInfo.Fields)
                        {
                            if (field.Name == "*")
                            {
                                fieldValues.Add(field.Name, jo.ToString());
                            }
                            else
                            {
                                fieldValues.Add(field.Name, "");
                            }
                        }
                    }
                    else if (jas[i] is JObject)
                    {
                        JObject jo = jas[i] as JObject;

                        foreach (Proj_Detail_Field field in detailPageInfo.Fields)
                        {
                            if (field.Name == "*")
                            {
                                fieldValues.Add(field.Name, jo.ToString());
                            }
                            else
                            {
                                JToken fieldValueJ = jo[field.Name];
                                string value = fieldValueJ == null ? "" : fieldValueJ.ToString();
                                fieldValues.Add(field.Name, value);
                            }
                        }
                    }

                    fieldValues.Add(SysConfig.DetailPageNameFieldName, pageName);
                    fieldValues.Add(SysConfig.DetailPageUrlFieldName, pageUrl);

                    filedValuesList.Add(fieldValues);
                }
                this.SaveDetailFieldValueToFile(filedValuesList, localReadFilePath);
            }
        }

        private bool GrabDetailPage(IListSheet listSheet, string pageUrl, Dictionary<string, string> listRow, string localPagePath, int pageIndex, Proj_Detail_MultiLine detailPageInfo, string cookie)
        {
            return GrabDetailPage(listSheet, pageUrl, listRow, localPagePath, pageIndex, detailPageInfo, cookie, "detail");
        }

        private bool GrabDetailPage(IListSheet listSheet, string pageUrl, Dictionary<string, string> listRow, string localPagePath, int pageIndex, Proj_Detail_MultiLine detailPageInfo, string cookie, string tabName)
        {
            string pageName = DetailPageNameList[pageIndex];
            decimal intervalAfterLoaded = detailPageInfo.IntervalAfterLoaded;
            Encoding encoding = Encoding.GetEncoding(detailPageInfo.Encoding);
            try
            {
                switch (detailPageInfo.DataAccessType)
                {
                    case Proj_DataAccessType.WebBrowserHtml:
                        {
                            string webPageText = GetDetailHtmlByWebBrowser(pageUrl, listRow, intervalAfterLoaded, detailPageInfo.RequestTimeout, detailPageInfo.CompleteChecks, tabName);
                            this.SaveFile(webPageText, localPagePath, encoding);
                        }
                        break;
                    case Proj_DataAccessType.WebRequestHtml:
                        {
                            string webPageText = GetTextByRequest(pageUrl, listRow, detailPageInfo.NeedProxy, intervalAfterLoaded, detailPageInfo.RequestTimeout, encoding, cookie, detailPageInfo.XRequestedWith, detailPageInfo.AutoAbandonDisableProxy, detailPageInfo.DataAccessType, detailPageInfo.CompleteChecks, detailPageInfo.IntervalProxyRequest);
                            this.SaveFile(webPageText, localPagePath, encoding);
                        }
                        break;
                    case Proj_DataAccessType.WebRequestJson:
                        {
                            string webPageText = GetTextByRequest(pageUrl, listRow, detailPageInfo.NeedProxy, intervalAfterLoaded, detailPageInfo.RequestTimeout, encoding, cookie, detailPageInfo.XRequestedWith, detailPageInfo.AutoAbandonDisableProxy, detailPageInfo.DataAccessType, detailPageInfo.CompleteChecks, detailPageInfo.IntervalProxyRequest);
                            this.SaveFile(webPageText, localPagePath, encoding);
                        }
                        break;
                    case Proj_DataAccessType.WebRequestFile:
                        {
                            byte[] data = GetFileByRequest(pageUrl, listRow, detailPageInfo.NeedProxy, intervalAfterLoaded, detailPageInfo.RequestTimeout, detailPageInfo.AutoAbandonDisableProxy, detailPageInfo.IntervalProxyRequest);
                            this.SaveFile(data, localPagePath);
                        }
                        break;
                }
                return true;
            }
            catch (NoneProxyException ex)
            {
                throw ex;
            }
            catch (Exception ex)
            {
                //当抓取了一个页面/文件出错后
                if (this.ExternalRunPage.AfterGrabOneCatchException(pageUrl, listRow, ex))
                {
                    string errorMsg = "线程" + Thread.CurrentThread.ManagedThreadId.ToString() + ": 放弃抓取（外部程序要求放弃）.";
                    listSheet.SetGiveUp(pageIndex, pageUrl, errorMsg);
                    this.InvokeAppendLogText(errorMsg + " PageUrl = " + pageUrl + ". " + ex.Message + (ex.InnerException == null ? "" : ex.InnerException.Message), LogLevelType.Error, true);
                    return true;     
                }
                else if (!detailPageInfo.AllowAutoGiveUp || !GiveUpGrabPage(listSheet, pageUrl, pageIndex, ex))
                {
                    this.InvokeAppendLogText("线程" + Thread.CurrentThread.ManagedThreadId.ToString() + ": PageUrl = " + pageUrl + ".\r\n" + ex.Message + (ex.InnerException == null ? "" : ex.InnerException.Message), LogLevelType.Error, true);
                    return false;
                }
                else
                {
                    this.InvokeAppendLogText("线程" + Thread.CurrentThread.ManagedThreadId.ToString() + ": 放弃抓取. PageUrl = " + pageUrl + ". \r\n" + ex.Message + (ex.InnerException == null ? "" : ex.InnerException.Message), LogLevelType.Error, true);
                    return true;
                }
            }
        }
        #endregion          

        #region 创建Excel文件,包含详情页地址
        public List<Proj_Detail_Field> GetOutputFields(List<Proj_Detail_Field> detailPageFields)
        {
            List<Proj_Detail_Field> newAllFields = new List<Proj_Detail_Field>();
            Proj_Detail_Field urlField = new Proj_Detail_Field();
            urlField.Name = SysConfig.DetailPageUrlFieldName;
            urlField.ColumnWidth = 20;
            newAllFields.Add(urlField); 
            if (detailPageFields != null)
            {
                newAllFields.AddRange(detailPageFields.ToArray());
            }
            return newAllFields;
        }
        #endregion

        #region 判断列表DB是否已经存在
        private bool NeedCreateListDB(string dbFilePath)
        {
            if (!this.MustReGrab && this.PopPrompt && File.Exists(dbFilePath))
            {
                if (MessageBox.Show("列表DB文件已经存在，是否继续执行上次未完成的任务?", "确认", MessageBoxButtons.OKCancel)
                    == DialogResult.OK)
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
            else
            {
                return true;
            }
        }
        #endregion 
        
        #region 将excel文件中的list复制到db中
        private IListSheet CopyListToDBFromExcel(string excelFilePath, string dbFilePath)
        {
            ExcelReader excelReader = new ExcelReader(excelFilePath, "List");

            //判断excel中必须的行是否存在
            Dictionary<string, int> columnToIndexs = excelReader.ColumnNameToIndex;
            List<string> excelSysColumnList = new List<string>();
            excelSysColumnList.Add(SysConfig.DetailPageUrlFieldName);
            excelSysColumnList.Add(SysConfig.DetailPageNameFieldName);
            excelSysColumnList.Add(SysConfig.DetailPageCookieFieldName);
            excelSysColumnList.Add(SysConfig.GrabStatusFieldName);
            excelSysColumnList.Add(SysConfig.GiveUpGrabFieldName);
            foreach (string columnName in excelSysColumnList)
            {
                if (!columnToIndexs.ContainsKey(columnName))
                {
                    throw new Exception("导入的详情页地址Excel中没有包含列: " + columnName);
                }
            }

            //获取Excel中的记录行
            List<Dictionary<string, string>> allExcelRows = new List<Dictionary<string, string>>();
            int excelRowCount = excelReader.GetRowCount();
            for (int i = 0; i < excelRowCount; i++)
            {
                Dictionary<string, string> f2vs = excelReader.GetFieldValues(i);
                string giveUpStr = f2vs[SysConfig.GiveUpGrabFieldName];
                if (CommonUtil.IsNullOrBlank(giveUpStr))
                {
                    f2vs[SysConfig.GiveUpGrabFieldName] = "N";
                }
                else if (giveUpStr == "是")
                {
                    f2vs[SysConfig.GiveUpGrabFieldName] = "Y";
                }
                f2vs[SysConfig.ListPageIndexFieldName] = i.ToString();
                allExcelRows.Add(f2vs);
            }
            excelReader.Close();

            //增加用户自定义列
            List<string> dbSysColumnList = new List<string>();
            dbSysColumnList.AddRange(excelSysColumnList.ToArray());
            dbSysColumnList.Add(SysConfig.ListPageIndexFieldName);

            IListSheet listSheet = new ListSheetInDB(dbFilePath);
            List<string> columns = new List<string>();
            foreach (string column in columnToIndexs.Keys)
            {
                if (!dbSysColumnList.Contains(column))
                {
                    columns.Add(column);
                }
            }
            listSheet.AddColumns(columns);

            //拷贝数据记录
            if (allExcelRows.Count > 0)
            {
                listSheet.CopyToListSheet(allExcelRows, columnToIndexs, excelSysColumnList);
            }
            return listSheet;
        }
        #endregion

        #region 将excel文件中的list复制到db中
        private IListSheet CopyListToDBFromExcel(string excelFilePath, IListSheet listSheet)
        {
            ExcelReader excelReader = new ExcelReader(excelFilePath, "List");

            //判断excel中必须的行是否存在
            Dictionary<string, int> columnToIndexs = excelReader.ColumnNameToIndex;
            List<string> excelSysColumnList = new List<string>();
            excelSysColumnList.Add(SysConfig.DetailPageUrlFieldName);
            excelSysColumnList.Add(SysConfig.DetailPageNameFieldName);
            excelSysColumnList.Add(SysConfig.DetailPageCookieFieldName);
            excelSysColumnList.Add(SysConfig.GrabStatusFieldName);
            excelSysColumnList.Add(SysConfig.GiveUpGrabFieldName);
            foreach (string columnName in excelSysColumnList)
            {
                if (!columnToIndexs.ContainsKey(columnName))
                {
                    throw new Exception("导入的详情页地址Excel中没有包含列: " + columnName);
                }
            }

            //获取Excel中的记录行
            List<Dictionary<string, string>> allExcelRows = new List<Dictionary<string, string>>();
            int beginPageIndex = listSheet.GetListDBRowCount();
            int excelRowCount = excelReader.GetRowCount();
            for (int i = 0; i < excelRowCount; i++)
            {
                Dictionary<string, string> f2vs = excelReader.GetFieldValues(i);
                string giveUpStr = f2vs[SysConfig.GiveUpGrabFieldName];
                if (CommonUtil.IsNullOrBlank(giveUpStr))
                {
                    f2vs[SysConfig.GiveUpGrabFieldName] = "N";
                }
                else if (giveUpStr == "是")
                {
                    f2vs[SysConfig.GiveUpGrabFieldName] = "Y";
                }
                f2vs[SysConfig.ListPageIndexFieldName] = (beginPageIndex + i).ToString();
                allExcelRows.Add(f2vs);
            }
            excelReader.Close(); 

            //拷贝数据记录
            if (allExcelRows.Count > 0)
            {
                listSheet.CopyToListSheet(allExcelRows, columnToIndexs, excelSysColumnList);
            }
            return listSheet;
        }
        #endregion

        #region 获取ListSheet连接
        private IListSheet GetDBListSheet(string dbFilePath)
        {
            IListSheet listSheet = new ListSheetInDB(dbFilePath);
            return listSheet;
        }
        #endregion

        #region 在List的DB文件中增加用户定义的项目中包含的列 
        private void InitCustomColumnsToListSheet(IListSheet listSheet)
        {
            List<string> fields = new List<string>(); 
            listSheet.AddColumns(fields);
        } 
        #endregion

        #region 创建DB文件
        public IListSheet CreateListSheet(string dbFilePath, string newListDBFilePath)
        {
            CommonUtil.CreateFileDirectory(dbFilePath);
            IListSheet listSheet = null;
            if (this.NeedCreateListDB(dbFilePath))
            {
                File.Copy(newListDBFilePath, dbFilePath, true);

                if (File.Exists(this.ListFilePath))
                {
                    InvokeAppendLogText("从Excel文件中读取下载地址...", LogLevelType.System, true);
                    listSheet = this.CopyListToDBFromExcel(this.ListFilePath, dbFilePath);
                    GC.Collect();
                }
                else if (File.Exists(ExcelFilePath))
                {
                    InvokeAppendLogText("从Excel文件中读取下载地址...", LogLevelType.System, true);
                    listSheet = this.CopyListToDBFromExcel(this.ExcelFilePath, dbFilePath); 
                    GC.Collect();
                }
                else if(File.Exists(this.GetPartExcelFilePath(1)))
                {
                    InvokeAppendLogText("从Excel文件中读取下载地址...", LogLevelType.System, true);
                    int fileIndex = 1; 
                    string excelFilePath = this.GetPartExcelFilePath(fileIndex);
                    if (File.Exists(excelFilePath))
                    {
                        InvokeAppendLogText("从Excel文件中读取下载地址, Part_" + fileIndex.ToString(), LogLevelType.System, true);
                        listSheet = this.CopyListToDBFromExcel(excelFilePath, dbFilePath);
                        GC.Collect();
                        while (1 == 1)
                        {
                            fileIndex++;
                            excelFilePath = this.GetPartExcelFilePath(fileIndex);
                            if (File.Exists(excelFilePath))
                            {
                                InvokeAppendLogText("从Excel文件中读取下载地址, Part_" + fileIndex.ToString(), LogLevelType.System, true);
                                this.CopyListToDBFromExcel(excelFilePath, listSheet);
                                GC.Collect();
                            }
                            else
                            {
                                break;
                            }
                        }
                    }
                } 
                else
                {
                    InvokeAppendLogText("无下载地址...", LogLevelType.System, true);
                    listSheet = this.GetDBListSheet(dbFilePath);
                    //增加用户自定义列
                    InitCustomColumnsToListSheet(listSheet);
                }
            }
            else
            {
                InvokeAppendLogText("读取历史导入的下载地址...", LogLevelType.System, true);
                listSheet = this.GetDBListSheet(dbFilePath);
            }

            return listSheet;
        }
        #endregion   

        #region 列名序号对应
        private Dictionary<string, int> _ColumnNameToIndex = null;
        public Dictionary<string, int> ColumnNameToIndex
        {
            get
            {
                if (_ColumnNameToIndex == null)
                {
                    List<Proj_Detail_Field> fields = new List<Proj_Detail_Field>();
                    Dictionary<string, int> columnNameToIndex = new Dictionary<string, int>();

                    //columnNameToIndex.Add(SysConfig.ListPageIndexFieldName, SysConfig.ListPageIndexFieldIndex);
                    columnNameToIndex.Add(SysConfig.DetailPageUrlFieldName, SysConfig.DetailPageUrlFieldIndex);
                    columnNameToIndex.Add(SysConfig.DetailPageNameFieldName, SysConfig.DetailPageNameFieldIndex);
                    columnNameToIndex.Add(SysConfig.DetailPageCookieFieldName, SysConfig.DetailPageCookieFieldIndex);
                    columnNameToIndex.Add(SysConfig.GrabStatusFieldName, SysConfig.GrabStatusFieldIndex);
                    columnNameToIndex.Add(SysConfig.GiveUpGrabFieldName, SysConfig.GiveUpGrabFieldIndex);

                    if (this.Project.DetailGrabType != DetailGrabType.NoneDetailPage)
                    {
                        fields.AddRange(this.Project.DetailGrabInfoObject.Fields.ToArray());
                    }

                    for (int i = 0; i < fields.Count; i++)
                    {
                        int index = i + SysConfig.SystemColumnCount;
                        Proj_Detail_Field field = fields[i];
                        columnNameToIndex.Add(field.Name, index);
                    }
                    _ColumnNameToIndex = columnNameToIndex;
                }
                return _ColumnNameToIndex;
            }
        }
        #endregion
         
        #region 设置列宽
        private void SetColumnWidth(ISheet sheet, int columnIndex, int width)
        {
            sheet.SetColumnWidth(columnIndex, width * 256);
        }
        #endregion

        #region 设置列宽
        private void SetCellValue(IRow row, int columnIndex, string value)
        {
            row.CreateCell(columnIndex).SetCellValue(value);
        }
        #endregion 

        #region 保存Excel文件到硬盘 放弃使用的方法
        /*
        public void SaveExcelToDisk( string filePath)
        { 
            CommonUtil.CreateFileDirectory(filePath);
             
            FileStream fs = null;
            try
            {
                fs = File.Open(filePath, FileMode.Create);
                //wk.Write(fs); 
            }
            catch (Exception ex)
            {
                throw new Exception("保存Excel文件到硬盘失败. FilePath = " + filePath, ex);
            }
            finally
            {
                if (fs != null)
                {
                    fs.Close();
                    fs.Dispose();
                }
            }
        }
         */
        #endregion

        #region 写入Excel文件
        public void SaveListFieldValueToDB(IListSheet sheet, Dictionary<string, string> fieldValues)
        {
            sheet.AddListRow(fieldValues); 
        }

        public void CheckPageUrlByExcel(ISheet listSheet, int rowIndex, string pageUrl)
        {
            IRow listRow = listSheet.GetRow(rowIndex + SysConfig.ColumnTitleRowCount);
            string urlCellValue = listRow.GetCell(SysConfig.DetailPageUrlFieldIndex).ToString();
            if (urlCellValue != pageUrl)
            { 
                throw new Exception("第" + rowIndex.ToString() + "行地址不匹配. Url_1 = " + pageUrl + ", Url_2 = " + urlCellValue);
            }
        }

        public void SaveDetailFieldValueToExcel(ISheet listSheet, ISheet detailSheet, Dictionary<string, string> fieldValues, int rowIndex, string pageUrl)
        {
            IRow listRow = listSheet.GetRow(rowIndex + SysConfig.ColumnTitleRowCount);
            string urlCellValue = listRow.GetCell(SysConfig.DetailPageUrlFieldIndex).ToString();
            if (urlCellValue == pageUrl)
            {
                IRow detailRow = detailSheet.CreateRow(detailSheet.LastRowNum + 1);
                foreach (string columnName in this.ColumnNameToIndex.Keys)
                {
                    int index = this.ColumnNameToIndex[columnName];
                    ICell cell = listRow.GetCell(index);
                    if (cell != null)
                    {
                        detailRow.CreateCell(index).SetCellValue(cell.ToString());
                    }
                }

                foreach (string fieldName in fieldValues.Keys)
                {
                    int index = this.ColumnNameToIndex[fieldName];
                    string value = fieldValues[fieldName];
                    ICell cell = detailRow.CreateCell(index);
                    cell.SetCellValue(value);
                }
            }
            else
            {
                throw new Exception("第" + rowIndex.ToString() + "行地址不匹配. Url_1 = " + pageUrl + ", Url_2 = " + urlCellValue);
            }
        }
        #endregion 

        #region 保存读取信息结果
        public void SaveDetailFieldValueToFile(Dictionary<string, string> fieldValues, string localReadFilePath)
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml("<?xml version=\"1.0\" encoding=\"UTF-8\"?><DetailPage></DetailPage>");
            XmlElement root = xmlDoc.DocumentElement;
            foreach (string f in fieldValues.Keys)
            {
                string value = fieldValues[f];
                XmlNode node = xmlDoc.CreateElement("Field");
                root.AppendChild(node);

                XmlAttribute fieldAttr = xmlDoc.CreateAttribute("Key");
                fieldAttr.Value = f;
                node.Attributes.Append(fieldAttr);

                XmlAttribute valueAttr = xmlDoc.CreateAttribute("Value");
                valueAttr.Value = value;
                node.Attributes.Append(valueAttr);
            }
            this.SaveFile(xmlDoc.OuterXml, localReadFilePath, Encoding.UTF8);
        }
        #endregion

        #region 读取信息结果
        public Dictionary<string, string> ReadDetailFieldValueFromFile(string localReadFilePath)
        {
            string xml = ReadFile(localReadFilePath);
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(xml);
            XmlElement root = xmlDoc.DocumentElement;
            Dictionary<string, string> fieldValues = new Dictionary<string, string>();
            XmlNodeList nodes = root.ChildNodes;

            foreach (XmlNode fNode in nodes)
            {
                string f = fNode.Attributes["Key"].Value;
                string v = fNode.Attributes["Value"].Value;
                fieldValues.Add(f, v);
            }
            return fieldValues;
        } 
        public List<Dictionary<string, string>> ReadDetailFieldValueListFromFile(string localReadFilePath)
        {
            string xml = ReadFile(localReadFilePath);
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(xml);
            XmlElement root = xmlDoc.DocumentElement;
            List<Dictionary<string, string>> fieldValueList = new List<Dictionary<string, string>>();
            XmlNodeList fvNodes = root.ChildNodes;

            foreach (XmlNode fvNode in fvNodes)
            {
                XmlNodeList fNodes = fvNode.ChildNodes;
                Dictionary<string, string> fieldValues = new Dictionary<string, string>();
                foreach (XmlNode fNode in fNodes)
                {
                    string f = fNode.Attributes["Key"].Value;
                    string v = fNode.Attributes["Value"].Value;
                    fieldValues.Add(f, v);
                }
                fieldValueList.Add(fieldValues);
            }
            return fieldValueList;
        }
        #endregion

        #region 保存读取信息结果
        public void SaveDetailFieldValueToFile(List<Dictionary<string, string>> fieldValuesList, string localReadFilePath)
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml("<?xml version=\"1.0\" encoding=\"UTF-8\"?><DetailPages></DetailPages>");
            XmlElement root = xmlDoc.DocumentElement;
            for (int i = 0; i < fieldValuesList.Count; i++)
            {
                Dictionary<string, string> fieldValues = fieldValuesList[i];
                XmlNode pNode = xmlDoc.CreateElement("DetailPage");
                root.AppendChild(pNode);

                foreach (string f in fieldValues.Keys)
                {
                    string value = fieldValues[f];
                    XmlNode node = xmlDoc.CreateElement("Field");
                    pNode.AppendChild(node);

                    XmlAttribute fieldAttr = xmlDoc.CreateAttribute("Key");
                    fieldAttr.Value = f;
                    node.Attributes.Append(fieldAttr);

                    XmlAttribute valueAttr = xmlDoc.CreateAttribute("Value");
                    valueAttr.Value = value;
                    node.Attributes.Append(valueAttr);
                }
            }
            this.SaveFile(xmlDoc.OuterXml, localReadFilePath ,Encoding.UTF8);
        }
        #endregion 

        #region 抓取进程
        private Thread _GrabThread;
        /// <summary>
        /// 抓取进程
        /// </summary>
        private Thread GrabThread
        {
            get
            {
                return _GrabThread;
            }
            set
            {
                _GrabThread = value;
            }
        }
        #endregion

        #region 显示网页
        public WebBrowser InvokeShowWebPage(string url, string tabName)
        {
            return (WebBrowser)this.Invoke(new ShowWebPageInvokeDelegate(this.ShowWebPage), new string[] { url, tabName });
        }
        private WebBrowser ShowWebPage(string url, string tabName)
        {
            NdaWebBrowser webBrowser = this.CreateWebBrowser(tabName);

            /*
            if (!Uri.IsWellFormedUriString(url, UriKind.Absolute))
            {
                Uri newUri = new Uri(webBrowser.Url, url);
                url = newUri.AbsoluteUri;
            }
            */
            IsCompleted[webBrowser.TabName] = false;
            webBrowser.Navigate(url);
            return webBrowser;
        }
        #endregion

        #region 关闭网页
        public void CloseWebPage(string tabName)
        {
            this.Invoke(new CloseWebPageInvokeDelegate(this.CloseWebPageMethod), new string[] { tabName });
        }
        private void CloseWebPageMethod(string tabName)
        {
            lock (tabLocker)
            {
                this.RemoveOldTabPageAndBrowser(tabName);
            }
        }
        #endregion

        #region 显示网页
        private delegate void ScrollWebPageInvokeDelegate(int toY, string tabName);
        private void InvokeScrollWebPage(int toY, string tabName)
        {
            this.Invoke(new ScrollWebPageInvokeDelegate(this.ScrollWebPage), new object[] { toY, tabName });
        }
        private void ScrollWebPage(int toY, string tabName)
        {
            WebBrowser webBrowser = this.GetWebBrowserByName(tabName);
            webBrowser.Document.Window.ScrollTo(0, toY);
        }
        #endregion

        #region 获取网页高度
        private delegate int GetWebPageHeightInvokeDelegate(string tabName);
        private int InvokeGetWebPageHeight(string tabName)
        {
            return (int)this.Invoke(new GetWebPageHeightInvokeDelegate(this.GetWebPageHeight), new object[] { tabName });
        }
        private int GetWebPageHeight(string tabName)
        {
            WebBrowser webBrowser = this.GetWebBrowserByName(tabName);
            return webBrowser.Document.Body.OffsetRectangle.Bottom;
        }
        #endregion

        #region 关闭
        public bool Close()
        {
            if (this.GrabThread != null && this.IsGrabing)
            {
                if (PopPrompt || CommonUtil.Confirm("提示", "正在执行抓取. 确定要关闭吗?"))
                {
                    try
                    {
                        this.GrabThread.Abort();
                        this.GrabThread = null;
                        if (this.AllGrabDetailThreads != null)
                        {
                            foreach (Thread thread in this.AllGrabDetailThreads)
                            {
                                if (thread != null && thread.IsAlive)
                                {
                                    thread.Abort();
                                }
                            }
                            this.AllGrabDetailThreads.Clear();
                        }
                    }
                    catch (Exception ex)
                    {
                        this.InvokeAppendLogText("关闭任务时发生错误. " + ex.Message, LogLevelType.Fatal, true);
                        if (!PopPrompt)
                        {
                            CommonUtil.Alert("错误", ex.Message);
                        }
                    }
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return true;
            }
        }
        #endregion 

        #region 开始抓取按钮
        
        private bool _AllowRunGrabDetail = true;
        private bool _AllowRunRead = true;
        private bool _AllowRunExport = true;
        private bool _AllowRunCustom = true;

        private void buttonBeginGrab_Click(object sender, EventArgs e)
        {
            DoGrab();
        }
        private void DoGrab()
        {
            if (this.ExternalRunPage.BeforeAllGrab())
            {
                if (!PopPrompt)
                {
                    BeginGrab();
                }
                else
                {
                    FormTaskSetting taskSettingForm = new FormTaskSetting();
                    if (taskSettingForm.ShowDialog() == DialogResult.OK)
                    {
                        Dictionary<string, object> setting = taskSettingForm.Setting;
                        this._AllowRunGrabDetail = (bool)setting["RunGrabDetail"];
                        this._AllowRunRead = (bool)setting["RunRead"];
                        this._AllowRunExport = (bool)setting["RunExport"];
                        this._AllowRunCustom = (bool)setting["RunCustom"];

                        BeginGrab();
                    }
                }
            }
        }  

        public void BeginGrab()
        {  
            Thread thread = new Thread(new ThreadStart(Grab));
            this.GrabThread = thread;
            IsGrabing = true;
            thread.Start();
        }
        #endregion

        #region 初始化之后
        protected void AfterLoad()
        {
            if (this.Project.LoginPageInfoObject != null)
            {
                Proj_LoginPageInfo loginObj = (Proj_LoginPageInfo)this.Project.LoginPageInfoObject;
                this.ShowWebPage(loginObj.LoginUrl, "login");
            }
        }
        #endregion 

        #region 给控件赋值
        private delegate void SetControlValueByIdInvokeDelegate(string id, string attributeName, string attributeValue, string tabName);
        public void InvokeSetControlValueById(string id, string attributeName, string attributeValue, string tabName)
        {
            this.Invoke(new SetControlValueByIdInvokeDelegate(SetControlValueById), new object[] { id, attributeName, attributeValue, tabName });
        }
        private void SetControlValueById(string id, string attributeName, string attributeValue, string tabName)
        {
            WebBrowser webBrowser = this.GetWebBrowserByName(tabName);
            HtmlElement element = webBrowser.Document.GetElementById(id);
            element.SetAttribute(attributeName, attributeValue);
        }
        #endregion

        #region 读取下载下来的html
        public HtmlAgilityPack.HtmlDocument GetLocalHtmlDocument(IListSheet listSheet, int pageIndex)
        {
            return GetLocalHtmlDocument(listSheet, pageIndex, Encoding.UTF8);
        }
        #endregion

        #region 读取下载下来的html
        public HtmlAgilityPack.HtmlDocument GetLocalHtmlDocument(IListSheet listSheet, int pageIndex, Encoding encoding)
        {
            string pageUrl = listSheet.PageUrlList[pageIndex];
            string pageSourceDir = this.GetDetailSourceFileDir();
            string localFilePath = this.GetFilePath(pageUrl, pageSourceDir);
            TextReader tr = null;

            try
            {
                tr = new StreamReader(localFilePath, encoding);
                string webPageHtml = tr.ReadToEnd();
                HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                htmlDoc.LoadHtml(webPageHtml);
                return htmlDoc;
            }
            catch (Exception ex)
            {
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
        }
        #endregion 

        #region 显示网页
        public WebBrowser ShowWebPage(string pageUrl, string tabName, int webRequestTimeout, bool goonWhenTimeout)
        {
            WebBrowser webBrowser = this.InvokeShowWebPage(pageUrl, tabName);
            int waitCount = 0;
            while (!this.CheckIsComplete(tabName))
            {
                if (SysConfig.WebPageRequestInterval * waitCount > webRequestTimeout)
                {
                    string errorInfo = "打开页面超时! PageUrl = " + pageUrl;
                    this.InvokeAppendLogText(errorInfo, Base.EnumTypes.LogLevelType.System, true);
                    if (!goonWhenTimeout)
                    {
                        throw new Exception(errorInfo);
                    }
                    else
                    {
                        break;
                    }
                }
                //等待
                waitCount++;
                Thread.Sleep(SysConfig.WebPageRequestInterval);
            }
            return webBrowser;
        }
        #endregion 

        #region 向网页中增加JavaScript代码，例如函数等，方便后续网页与爬取工具交互
        public void InvokeAddScriptMethod(WebBrowser webBrowser, string scriptMethodCode, object objectForScripting)
        {
            webBrowser.Invoke(new AddScriptMethodDelegate(AddScriptMethod), new object[] { webBrowser, scriptMethodCode , objectForScripting});
        }
        private delegate void AddScriptMethodDelegate(WebBrowser webBrowser, string scriptMethodCode, object objectForScripting);
        private void AddScriptMethod(WebBrowser webBrowser, string scriptMethodCode, object objectForScripting)
        {
            if (objectForScripting == null)
            {
                webBrowser.ObjectForScripting = this;
            }
            else
            {
                webBrowser.ObjectForScripting = objectForScripting;
            }
            HtmlElement sElement = webBrowser.Document.CreateElement("script");
            IHTMLScriptElement scriptElement = (IHTMLScriptElement)sElement.DomElement;
            scriptElement.text = scriptMethodCode;
            webBrowser.Document.Body.AppendChild(sElement);
        }
        #endregion

        #region 调用网页JavaScript脚本
        public object InvokeDoScriptMethod(WebBrowser webBrowser, string methodName, object[] parameters)
        {
            return webBrowser.Invoke(new DoScriptMethodDelegate(DoScriptMethod), new object[] { webBrowser, methodName, parameters });
        }
        private delegate object DoScriptMethodDelegate(WebBrowser webBrowser, string methodName, object[] parameters);
        private object DoScriptMethod(WebBrowser webBrowser, string methodName, object[] parameters)
        {
            return webBrowser.Document.InvokeScript(methodName, parameters);
        }
        #endregion

        #region 滚动页面
        public void InvokeScrollDocumentMethod(WebBrowser webBrowser, Point toPoint)
        {
            webBrowser.Invoke(new ScrollDocumentMethodDelegate(ScrollDocumentMethod), new object[] { webBrowser, toPoint });
        }
        private delegate void ScrollDocumentMethodDelegate(WebBrowser webBrowser, Point toPoint);
        private void ScrollDocumentMethod(WebBrowser webBrowser, Point toPoint)
        {
            webBrowser.Document.Window.ScrollTo(toPoint);
        }
        #endregion

        #region 滚动页面
        public void InvokeWebBrowserGoBackMethod(WebBrowser webBrowser)
        {
            webBrowser.Invoke(new WebBrowserGoBackMethodDelegate(WebBrowserGoBackMethod), new object[] { webBrowser });
        }
        private delegate void WebBrowserGoBackMethodDelegate(WebBrowser webBrowser);
        private void WebBrowserGoBackMethod(WebBrowser webBrowser)
        {
            webBrowser.GoBack();
        }
        #endregion  
         
        #region 实现了轮询的方法判断网页JavaScript里某个值是否等于checkValue。用于异步调用后等待执行结果
        public void WaitForInvokeScript(WebBrowser webBrowser, string scriptCheckMethod, string checkValue, int invokeTimeout)
        {
            //记录调用check方法返回的值
            string resultValue = null;

            int waitCount = 0;

            //存在异步加载数据的情况，此处用轮询获取查询到的数据
            while (resultValue == null)
            {
                if (SysConfig.WebPageRequestInterval * waitCount > invokeTimeout)
                {
                    //超时
                    string errorInfo = "执行" + scriptCheckMethod + "方法获取返回值超时";
                    this.InvokeAppendLogText(errorInfo, LogLevelType.Error, true);
                    throw new Exception(errorInfo);
                }
                Thread.Sleep(SysConfig.WebPageRequestInterval);
                waitCount++;
                resultValue = (string)InvokeDoScriptMethod(webBrowser, scriptCheckMethod, null);
            }
        }
        #endregion

        #region 利用成功和失败关键词判断是否打开了需要的页面
        public bool CheckOpenRightPage(WebBrowser webBrowser, string[] rightStrings, string[] wrongStrings, int timeout, bool andCondition)
        {
            bool matchRight = false;
            bool matchWrong = false;
            int waitCount = 0;

            while (!matchRight && !matchWrong)
            {
                if (SysConfig.WebPageRequestInterval * waitCount > timeout)
                {
                    //超时
                    string errorInfo = "网页加载失败, 无法判断是否打开了需要的页面, rightStrings=" + CommonUtil.StringArrayToString(rightStrings, ", ") + ", rightStrings=" + CommonUtil.StringArrayToString(rightStrings, ", ");
                    this.InvokeAppendLogText(errorInfo, LogLevelType.Error, true);
                    throw new Exception(errorInfo);
                }
                Thread.Sleep(SysConfig.WebPageRequestInterval);
                waitCount++;
                matchRight = InvokeCheckWebBrowserContains(webBrowser, rightStrings, andCondition);
                matchWrong = InvokeCheckWebBrowserContains(webBrowser, wrongStrings, andCondition);
            }
            return matchRight;
        }
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
        public bool CheckWebBrowserContainsForComplete(WebBrowser webBrowser, string[] checkStrings, int timeout, bool andCondition)
        {
            bool match = false;
            int waitCount = 0;

            while (!match)
            {
                if (SysConfig.WebPageRequestInterval * waitCount > timeout)
                {
                    //超时
                    string errorInfo = "网页加载失败, 无法匹配" + CommonUtil.StringArrayToString(checkStrings, ", ");
                    this.InvokeAppendLogText(errorInfo, LogLevelType.Error, true);
                    throw new Exception(errorInfo);
                }
                Thread.Sleep(SysConfig.WebPageRequestInterval);
                waitCount++;
                match = InvokeCheckWebBrowserContains(webBrowser, checkStrings, andCondition);
            }
            return match;
        } 

        public bool InvokeCheckWebBrowserContains(WebBrowser webBrowser, string[] checkStrings, bool andCondition)
        {
            return (bool)webBrowser.Invoke(new CheckWebBrowserContainsMethodDelegate(CheckWebBrowserContainsMethod), new object[] { webBrowser, checkStrings, andCondition });
        }
        private delegate bool CheckWebBrowserContainsMethodDelegate(WebBrowser webBrowser, string[] checkStrings, bool andCondition);
        private bool CheckWebBrowserContainsMethod(WebBrowser webBrowser, string[] checkStrings, bool andCondition)
        {
            string html = webBrowser.Document == null ? null : (webBrowser.Document.Body == null ? null : webBrowser.Document.Body.OuterHtml);
            if (html != null)
            {
                if (andCondition)
                {
                    foreach (string checkStr in checkStrings)
                    {
                        if (checkStr != null)
                        {

                            if (!html.Contains(checkStr))
                            {
                                return false;
                            }
                        }
                    }
                    return true;
                }
                else
                {
                    foreach (string checkStr in checkStrings)
                    {
                        if (checkStr != null)
                        {
                            if (html.Contains(checkStr))
                            {
                                return true;
                            }
                        }
                    }
                }
            }
            return false;
        }
        #endregion

        #region 判断浏览器是否已经跳转到某个网页
        public bool CheckWebBrowserUrl(WebBrowser webBrowser, string checkUrl, bool fullMatch, int timeout)
        {
            bool match = false;
            int waitCount = 0;

            //存在异步加载数据的情况，此处用轮询获取查询到的数据
            while (!match)
            {
                if (SysConfig.WebPageRequestInterval * waitCount > timeout)
                {
                    //超时
                    string errorInfo = "网页加载失败, 无法匹配" + checkUrl;
                    this.InvokeAppendLogText(errorInfo, LogLevelType.Error, true);
                    throw new Exception(errorInfo);
                }
                Thread.Sleep(SysConfig.WebPageRequestInterval);
                waitCount++;
                string pageUrl = InvokeGetWebBrowserPageUrl(webBrowser);
                match = fullMatch ? (pageUrl == checkUrl) : pageUrl.Contains(checkUrl);
            }
            return match;
        }

        public string InvokeGetWebBrowserPageUrl(WebBrowser webBrowser) 
        {
            return (string)webBrowser.Invoke(new GetWebBrowserPageUrlMethodDelegate(GetWebBrowserPageUrlMethod), new object[] { webBrowser });
        }
        private delegate string GetWebBrowserPageUrlMethodDelegate(WebBrowser webBrowser);
        private string GetWebBrowserPageUrlMethod(WebBrowser webBrowser)
        {
            return webBrowser.Url.AbsoluteUri;
        }
        #endregion

        #region 从中间文件读取信息
        public List<string> TryGetInfoFromMiddleFile(string fileName, string fieldName)
        {
            string localFilePath = this.GetFilePath(fileName, this.GetDetailSourceFileDir());
            CsvReader reader = CsvReader.TryLoad(localFilePath);
            if (reader == null)
            {
                return null;
            }
            else
            {
                List<String> allValues = reader.GetColumnValues(fieldName);
                return allValues;
            }
        }
        #endregion

        #region 将信息写入到中间文件
        public void SaveInfoToMiddleFile(string fileName, string fieldName, List<string> values)
        {
            string localFilePath = this.GetFilePath(fileName, this.GetDetailSourceFileDir());
            Dictionary<string, int> columnToIndex = new Dictionary<string,int>();
            columnToIndex.Add(fieldName, 0);
            CsvWriter writer = new CsvWriter(localFilePath, columnToIndex);
            writer.AddRows(fieldName, values);
            writer.SaveToDisk();
        }
        #endregion

        #region 从中间文件读取信息
        public List<Dictionary<string, string>> TryGetInfoFromMiddleFile(string fileName, string[] fieldNames)
        {
            string localFilePath = this.GetFilePath(fileName, this.GetDetailSourceFileDir());
            CsvReader reader = CsvReader.TryLoad(localFilePath);
            if (reader == null)
            {
                return null;
            }
            else
            {
                List<Dictionary<string, string>> allValues = reader.GetColumnValues(fieldNames);
                return allValues;
            }
        }
        #endregion

        #region 将信息写入到中间文件
        public void SaveInfoToMiddleFile(string fileName, string[] fieldNames, List<Dictionary<string, string>> valuesList)
        {
            string localFilePath = this.GetFilePath(fileName, this.GetDetailSourceFileDir());
            Dictionary<string, int> columnToIndex = new Dictionary<string, int>();
            for (int i = 0; i < fieldNames.Length; i++)
            {
                columnToIndex.Add(fieldNames[i], i);
            }
            CsvWriter writer = new CsvWriter(localFilePath, columnToIndex);
            writer.AddRows(valuesList);
            writer.SaveToDisk();
        }
        #endregion
    }
}
