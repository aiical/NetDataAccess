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
using NetDataAccess.Base.Web;
using System.Net;

namespace NetDataAccess.Extended.Anjuke
{
    public class GetCityAreaXiaoquList : ExternalRunWebPage
    {
        public override void WebRequestHtml_BeforeSendRequest(string pageUrl, Dictionary<string, string> listRow, Base.Web.NDAWebClient client)
        {
            string userAgent = this.RunPage.CurrentUserAgents.GetOneUserAgent();
            client.Headers["User-Agent"] = userAgent;
        }
        public override void CheckRequestCompleteFile(string webPageText, Dictionary<string, string> listRow)
        { 
            if (webPageText.Contains("msg") && webPageText.Contains("data"))
            {
            }
            else
            {
                throw new Exception("未完全加载文件.");
            }
        }

        private string GetListRowGrabCompleteMarkFilePath(string pageUrl, string sourceDir)
        {
            string localPagePath = this.RunPage.GetFilePath(pageUrl, sourceDir);
            return localPagePath;
        }

        public override bool CheckNeedGrab(Dictionary<string, string> listRow, string localPagePath)
        {
            return !File.Exists(localPagePath);
        }

        private int InitGrabDetailPageIndexList(IListSheet listSheet, string sourceDir)
        {
            int detailPageIndex = 0;
            this.RunPage.NeedGrabIndexs = new List<int>();
            this.RunPage.InvokeAppendLogText("开始统计需要下载的页面.", LogLevelType.System, true);
            while (detailPageIndex < this.RunPage.DetailPageUrlList.Count)
            {
                string pageUrl = this.RunPage.DetailPageUrlList[detailPageIndex];
                string localPagePath = this.RunPage.GetFilePath(pageUrl, sourceDir);
                if (this.CheckNeedGrab(listSheet.GetRow(detailPageIndex), localPagePath)
                    && !this.RunPage.CheckGiveUpGrabPage(listSheet, pageUrl, detailPageIndex))
                {
                    this.RunPage.NeedGrabIndexs.Add(detailPageIndex);
                }
                detailPageIndex++;
                if (detailPageIndex % 1000 == 0)
                {
                    double perc = (double)detailPageIndex / (double)this.RunPage.DetailPageUrlList.Count;
                    this.RunPage.InvokeAppendLogText("正在统计需要下载的页面..." + perc.ToString("#0.00%"), LogLevelType.System, true);
                }
            }
            this.RunPage.InvokeAppendLogText("完成统计需要下载的页面.", LogLevelType.System, true);
            return this.RunPage.NeedGrabIndexs.Count;
        }

        public override bool BeginGrabDetailPageInExternalProgram(IListSheet listSheet, Proj_Detail_SingleLine detailPageInfo)
        {
            string sourceDir = this.RunPage.GetSourceFileDir(detailPageInfo);
            this.RunPage.AllNeedGrabCount = this.InitGrabDetailPageIndexList(listSheet, sourceDir);
            if (this.RunPage.AllNeedGrabCount != 0)
            {
                this.BeginGrabDetailPageInParallelThread(listSheet, detailPageInfo);
                while (this.RunPage.CompleteGrabCount < this.RunPage.AllNeedGrabCount && this.RunPage.AllGrabDetailThreads.Count > 0)
                {
                    Thread.Sleep(5000);
                }
                return this.RunPage.SucceedGrabCount == this.RunPage.AllNeedGrabCount;
            }
            else
            {
                return true;
            }
        }

        private void BeginGrabDetailPageInParallelThread(IListSheet listSheet, Proj_Detail_SingleLine detailPageInfo)
        {
            //int threadCount = detailPageInfo.DataAccessType == Proj_DataAccessType.WebBrowserHtml ? 1 : detailPageInfo.ThreadCount;
            int threadCount = detailPageInfo.ThreadCount;
            this.RunPage.AllGrabDetailThreads = new List<Thread>();
            for (int i = 0; i < threadCount; i++)
            {
                Thread grabThread = new Thread(new ParameterizedThreadStart(ThreadGrabDetailPage));
                this.RunPage.AllGrabDetailThreads.Add(grabThread);
                this.RunPage.InvokeAppendLogText("线程" + grabThread.ManagedThreadId.ToString() + "开始抓取数据.", LogLevelType.System, true);
                grabThread.Start(new object[] { listSheet, detailPageInfo });
                Thread.Sleep(50);
            }
        }

        private void ThreadGrabDetailPage(object parameters)
        {
            object[] parameterArray = (object[])parameters;
            IListSheet listSheet = (IListSheet)parameterArray[0];
            Proj_Detail_SingleLine detailPageInfo = (Proj_Detail_SingleLine)parameterArray[1];
            string sourceDir = this.RunPage.GetSourceFileDir(detailPageInfo);
            Nullable<int> nextPageIndex = this.RunPage.GetNextGrabDetailPageIndex();
            while (nextPageIndex != null)
            {
                try
                {
                    this.ThreadGrabDetailPage(listSheet, (int)nextPageIndex, detailPageInfo, sourceDir);
                }
                catch (NoneProxyException ex)
                {
                    this.RunPage.InvokeAppendLogText("线程" + Thread.CurrentThread.ManagedThreadId.ToString() + ": 停止线程\r\n\r\n\r\n\r\n\r\n." + ex.Message, LogLevelType.System, true);
                    break;
                }
                catch (Exception ex)
                {
                    this.RunPage.InvokeAppendLogText("线程" + Thread.CurrentThread.ManagedThreadId.ToString() + ": 出错!!!!!!!!!!!!" + ex.Message, LogLevelType.System, true);
                }
                nextPageIndex = this.RunPage.GetNextGrabDetailPageIndex();
            }
            this.RunPage.AllGrabDetailThreads.Remove(Thread.CurrentThread);
        }

        private void ThreadGrabDetailPage(IListSheet listSheet, int detailPageIndex, Proj_Detail_SingleLine detailPageInfo, string sourceDir)
        {
            DateTime dt1 = DateTime.Now;
            string pageUrl = this.RunPage.DetailPageUrlList[detailPageIndex];
            string cookie = this.RunPage.DetailPageCookieList[detailPageIndex];
            string localPagePath = this.RunPage.GetFilePath(pageUrl, sourceDir);
            Dictionary<string, string> listRow = listSheet.GetRow(detailPageIndex);

            bool succeed = true;
            bool existLocalFile = File.Exists(localPagePath);


            if (!existLocalFile)
            {
                succeed = this.GrabDetailPage(listSheet, pageUrl, listRow, localPagePath, detailPageIndex, detailPageInfo, cookie);
            }

            this.RunPage.RefreshGrabCount(succeed);

            DateTime dt2 = DateTime.Now;
            TimeSpan ts = dt2 - dt1;
            this.RunPage.InvokeAppendLogText("线程" + Thread.CurrentThread.ManagedThreadId.ToString() + ": 抓取了第" + (detailPageIndex + 1).ToString() + "个页面, 用时" + ts.TotalSeconds.ToString("0.00") + "秒", LogLevelType.Normal, false);

            this.RunPage.RecordGrabDetailStatus(succeed, dt1, dt2);
        }

        private bool GrabDetailPage(IListSheet listSheet, string pageUrl, Dictionary<string, string> listRow, string localPagePath, int pageIndex, Proj_Detail_SingleLine detailPageInfo, string cookie)
        {
            string tabName = Thread.CurrentThread.ManagedThreadId.ToString();
            return this.GrabDetailPage(listSheet, pageUrl, listRow, localPagePath, pageIndex, detailPageInfo, cookie, tabName);
        }

        private bool GrabDetailPage(IListSheet listSheet, string pageUrl, Dictionary<string, string> listRow, string localPagePath, int pageIndex, Proj_Detail_SingleLine detailPageInfo, string cookie, string tabName)
        {
            string pageName = this.RunPage.DetailPageNameList[pageIndex];
            decimal intervalAfterLoaded = detailPageInfo.IntervalAfterLoaded;
            Encoding encoding = Encoding.GetEncoding(detailPageInfo.Encoding);
            string lastWebPageText = "";
            try
            {
                bool gotLastPage = false;
                int requestPageIndex = 1;
                while (!gotLastPage)
                {
                    string indexPageUrl = pageUrl + "?p=" + requestPageIndex.ToString();
                    string indexPageFilePath = this.RunPage.GetFilePath(indexPageUrl, this.RunPage.GetSourceFileDir(detailPageInfo));
                    if (!File.Exists(indexPageFilePath))
                    {
                        string webPageText = GetTextByRequest(indexPageUrl, listRow, detailPageInfo.NeedProxy, intervalAfterLoaded, detailPageInfo.RequestTimeout, encoding, cookie, detailPageInfo.XRequestedWith, detailPageInfo.AutoAbandonDisableProxy, detailPageInfo.DataAccessType, detailPageInfo.CompleteChecks, detailPageInfo.IntervalProxyRequest);
                        if (webPageText.Contains("\"msg\":\"ok\""))
                        {
                            if (webPageText.Contains("\"data\":[]") || lastWebPageText == webPageText)
                            {
                                //已到达最后一页
                                break;
                            }
                            else
                            {
                                this.RunPage.SaveFile(webPageText, indexPageFilePath, encoding);
                                lastWebPageText = webPageText;
                            }
                        }
                        else
                        {
                            throw new Exception("抓取出错: " + webPageText);
                        }
                    }
                    requestPageIndex++;
                }
                this.RunPage.SaveFile((requestPageIndex - 1).ToString(), localPagePath, encoding);

                return true;
            }
            catch (NoneProxyException ex)
            {
                throw ex;
            }
            catch (Exception ex)
            {
                if (!detailPageInfo.AllowAutoGiveUp || !this.RunPage.GiveUpGrabPage(listSheet, pageUrl, pageIndex, ex))
                {
                    this.RunPage.InvokeAppendLogText("线程" + Thread.CurrentThread.ManagedThreadId.ToString() + ": PageUrl = " + pageUrl + ". " + ex.Message + (ex.InnerException == null ? "" : ex.InnerException.Message), LogLevelType.Error, true);
                    return false;
                }
                else
                {
                    this.RunPage.InvokeAppendLogText("线程" + Thread.CurrentThread.ManagedThreadId.ToString() + ": 放弃抓取. PageUrl = " + pageUrl + ". " + ex.Message + (ex.InnerException == null ? "" : ex.InnerException.Message), LogLevelType.Error, true);
                    return true;
                }
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
                    client.ProxyServer = this.RunPage.CurrentProxyServers.BeginUse(intervalProxyRequest);
                }
                //client.Headers.Add("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; .NET CLR 1.0.3705;)");

                string userAgent = this.RunPage.CurrentUserAgents.GetOneUserAgent();
                client.Headers.Add("user-agent", userAgent);
                if (!CommonUtil.IsNullOrBlank(cookie))
                {
                    client.Headers.Add("cookie", cookie);
                    //client.Headers.Add("connection", "keep-alive");
                }
                client.Headers.Add("x-requested-with", "XMLHttpRequest");  

                client.OpenReadCompleted += client_OpenReadCompleted;
                client.OpenReadAsync(new Uri(pageUrl));

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

                        this.CheckRequestCompleteFile(s, listRow);

                        if (needProxy)
                        {
                            this.RunPage.CurrentProxyServers.Success(client.ProxyServer);
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
                        this.RunPage.CurrentProxyServers.Error(client.ProxyServer);
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
                    this.RunPage.CurrentProxyServers.EndUse(client.ProxyServer);
                }
            }
        }

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

        private void client_OpenReadCompleted(object sender, OpenReadCompletedEventArgs e)
        {
            if (this.RunPage.Grabing)
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
                    this.RunPage.InvokeAppendLogText("ReadToEnd获取字符串超时. " + ex.Message, LogLevelType.System, true);
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

        public override bool AfterAllGrab(IListSheet listSheet)
        {
            try
            {
                //this.GetCityXiaoquList(listSheet);
                //this.GetCityXiaoquDetailPageUrls(listSheet);
                this.GetXiaoquErshoufangListPageUrls(listSheet);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return true;
        } 

        private void GetCityXiaoquList(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>(); 
            resultColumnDic.Add("cityCode", 0);
            resultColumnDic.Add("cityName", 1);
            resultColumnDic.Add("level1AreaCode", 2);
            resultColumnDic.Add("level1AreaName", 3);
            resultColumnDic.Add("level2AreaCode", 4);
            resultColumnDic.Add("level2AreaName", 5);
            resultColumnDic.Add("name", 6);
            resultColumnDic.Add("address", 7);
            resultColumnDic.Add("sale_num", 8);
            resultColumnDic.Add("build_year", 9);
            resultColumnDic.Add("mid_price", 10);
            resultColumnDic.Add("url", 11);

            string resultFilePath = Path.Combine(exportDir, "安居客小区列表.csv");
            CsvWriter resultEW = new CsvWriter(resultFilePath, resultColumnDic);
            Dictionary<string, string> urlDic = new Dictionary<string, string>();

            

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                if (i % 100 == 0)
                {
                    this.RunPage.InvokeAppendLogText("正在输出CSV文件... " + ((double)(i * 100) / (double)listSheet.RowCount).ToString("0.00") + "%", LogLevelType.System, true);
                }

                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"];
                string cityName = row["cityName"];
                string cityCode = row["cityCode"];
                string level1AreaName = row["level1AreaName"];
                string level1AreaCode = row["level1AreaCode"];
                string level2AreaCode = row["level2AreaCode"];
                string level2AreaName = row["level2AreaName"];
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string localFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir);
                    string fileText = FileHelper.GetTextFromFile(localFilePath);
                    int requestPageCount = int.Parse(fileText);
                    for (int j = 0; j < requestPageCount; j++)
                    {
                        string pageLocalFilePath = this.RunPage.GetFilePath(detailUrl + "?p=" + (j + 1).ToString(), pageSourceDir);
                        string pageFileText = FileHelper.GetTextFromFile(pageLocalFilePath);
                        try
                        {
                            JObject rootJo = JObject.Parse(pageFileText);
                            JArray xiaoquJsonArray = rootJo["data"] as JArray;
                            for (int k = 0; k < xiaoquJsonArray.Count; k++)
                            {
                                JObject xiaoquJson = xiaoquJsonArray[k] as JObject;
                                string name = CommonUtil.HtmlDecode(xiaoquJson["name"].ToString());
                                string area = CommonUtil.HtmlDecode(xiaoquJson["area"].ToString());
                                string address = CommonUtil.HtmlDecode(xiaoquJson["address"].ToString());
                                string sale_num = xiaoquJson["sale_num"].ToString();
                                string build_year = CommonUtil.HtmlDecode(xiaoquJson["build_year"].ToString());
                                string mid_price = xiaoquJson["mid_price"].ToString();
                                string url = CommonUtil.HtmlDecode(xiaoquJson["url"].ToString());
                                if (!urlDic.ContainsKey(url))
                                {
                                    urlDic.Add(url, null);
                                    Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                    //f2vs.Add("detailPageUrl", url);
                                    //f2vs.Add("detailPageName", url);
                                    f2vs.Add("cityCode", cityCode);
                                    f2vs.Add("cityName", cityName);
                                    f2vs.Add("level1AreaCode", cityCode);
                                    f2vs.Add("level1AreaName", cityName);
                                    f2vs.Add("level2AreaCode", cityCode);
                                    f2vs.Add("level2AreaName", cityName);
                                    f2vs.Add("name", name);
                                    f2vs.Add("address", address);
                                    f2vs.Add("sale_num", sale_num);
                                    f2vs.Add("build_year", build_year);
                                    f2vs.Add("mid_price", mid_price);
                                    f2vs.Add("url", url);
                                    resultEW.AddRow(f2vs);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            throw ex;
                        }
                    }
                }
            }
            resultEW.SaveToDisk();
            this.RunPage.InvokeAppendLogText("完成输出CSV文件... 100%", LogLevelType.System, true);
        }

        private void GetCityXiaoquDetailPageUrls(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("cityCode", 5);
            resultColumnDic.Add("cityName", 6);
            resultColumnDic.Add("level1AreaCode", 7);
            resultColumnDic.Add("level1AreaName", 8);
            resultColumnDic.Add("level2AreaCode", 9);
            resultColumnDic.Add("level2AreaName", 10);
            resultColumnDic.Add("name", 11);
            resultColumnDic.Add("address", 12);
            resultColumnDic.Add("sale_num", 13);
            resultColumnDic.Add("build_year", 14);
            resultColumnDic.Add("mid_price", 15); 

            string resultFilePath = Path.Combine(exportDir, "安居客小区详情页.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic);
            Dictionary<string, string> urlDic = new Dictionary<string, string>();

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                if (i % 100 == 0)
                {
                    this.RunPage.InvokeAppendLogText("正在输出Excel文件... " + ((double)(i * 100) / (double)listSheet.RowCount).ToString("0.00") + "%", LogLevelType.System, true);
                }
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"];
                string cityName = row["cityName"];
                string cityCode = row["cityCode"];
                string level1AreaName = row["level1AreaName"];
                string level1AreaCode = row["level1AreaCode"];
                string level2AreaCode = row["level2AreaCode"];
                string level2AreaName = row["level2AreaName"];
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string localFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir);
                    string fileText = FileHelper.GetTextFromFile(localFilePath);
                    int requestPageCount = int.Parse(fileText);
                    for (int j = 0; j < requestPageCount; j++)
                    {
                        string pageLocalFilePath = this.RunPage.GetFilePath(detailUrl + "?p=" + (j + 1).ToString(), pageSourceDir);
                        string pageFileText = FileHelper.GetTextFromFile(pageLocalFilePath);
                        try
                        {
                            JObject rootJo = JObject.Parse(pageFileText);
                            JArray xiaoquJsonArray = rootJo["data"] as JArray;
                            for (int k = 0; k < xiaoquJsonArray.Count; k++)
                            {
                                JObject xiaoquJson = xiaoquJsonArray[k] as JObject;
                                string name = CommonUtil.HtmlDecode(xiaoquJson["name"].ToString());
                                string area = CommonUtil.HtmlDecode(xiaoquJson["area"].ToString());
                                string address = CommonUtil.HtmlDecode(xiaoquJson["address"].ToString());
                                string sale_num = xiaoquJson["sale_num"].ToString();
                                string build_year = CommonUtil.HtmlDecode(xiaoquJson["build_year"].ToString());
                                string mid_price = xiaoquJson["mid_price"].ToString();
                                string url = CommonUtil.HtmlDecode(xiaoquJson["url"].ToString());
                                if (!urlDic.ContainsKey(url))
                                {
                                    urlDic.Add(url, null);
                                    Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                    f2vs.Add("detailPageUrl", url);
                                    f2vs.Add("detailPageName", url);
                                    f2vs.Add("cityCode", cityCode);
                                    f2vs.Add("cityName", cityName);
                                    f2vs.Add("level1AreaCode", level1AreaCode);
                                    f2vs.Add("level1AreaName", level1AreaName);
                                    f2vs.Add("level2AreaCode", level2AreaCode);
                                    f2vs.Add("level2AreaName", level2AreaName);
                                    f2vs.Add("name", name);
                                    f2vs.Add("address", address);
                                    f2vs.Add("sale_num", sale_num);
                                    f2vs.Add("build_year", build_year);
                                    f2vs.Add("mid_price", mid_price); 
                                    resultEW.AddRow(f2vs);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            throw ex;
                        }
                    }
                }
            }
            resultEW.SaveToDisk();
            this.RunPage.InvokeAppendLogText("完成输出Excel文件... 100%", LogLevelType.System, true);
        }
        private ExcelWriter GetErshoufangExcelWriter(int fileIndex, string cityName)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("xiaoquUrl", 5);
            resultColumnDic.Add("xiaoquName", 6);
            resultColumnDic.Add("saleNum", 7);
            resultColumnDic.Add("cityName", 8);
            resultColumnDic.Add("cityCode", 9);
            resultColumnDic.Add("level1AreaName", 10);
            resultColumnDic.Add("level1AreaCode", 11);
            resultColumnDic.Add("level2AreaCode", 12);
            resultColumnDic.Add("level2AreaName", 13);

            string resultFilePath = Path.Combine(exportDir, "安居客小区二手房列表页_" + cityName + "_" + fileIndex.ToString() + ".xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }

        private void GetXiaoquErshoufangListPageUrls(IListSheet listSheet)
        {
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            Dictionary<string, ExcelWriter> cityToEW = new Dictionary<string, ExcelWriter>(); 
            Dictionary<string, int > cityToFileIndex = new Dictionary<string,int>(); 
            Dictionary<string, string> urlDic = new Dictionary<string, string>();

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"];
                string cityName = row["cityName"];
                ExcelWriter resultEW = cityToEW.ContainsKey(cityName) ? cityToEW[cityName] : null;

                if (resultEW == null || resultEW.RowCount > 500000)
                {
                    if (resultEW != null)
                    {
                        resultEW.SaveToDisk();
                    }
                    int fileIndex = cityToFileIndex.ContainsKey(cityName) ? cityToFileIndex[cityName] : 1;
                    resultEW = this.GetErshoufangExcelWriter(fileIndex,cityName);
                    fileIndex++;
                    cityToEW[cityName] = resultEW;
                    cityToFileIndex[cityName] = fileIndex;
                }

                if (i % 100 == 0)
                {
                    this.RunPage.InvokeAppendLogText("正在输出二手房列表页Excel文件... " + ((double)(i * 100) / (double)listSheet.RowCount).ToString("0.00") + "%", LogLevelType.System, true);
                }

                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string localFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir);
                    string fileText = FileHelper.GetTextFromFile(localFilePath);
                    int requestPageCount = int.Parse(fileText);
                    for (int j = 0; j < requestPageCount; j++)
                    {
                        string pageLocalFilePath = this.RunPage.GetFilePath(detailUrl + "?p=" + (j + 1).ToString(), pageSourceDir);
                        string pageFileText = FileHelper.GetTextFromFile(pageLocalFilePath);
                        try
                        {
                            JObject rootJo = JObject.Parse(pageFileText);
                            JArray xiaoquJsonArray = rootJo["data"] as JArray;
                            for (int k = 0; k < xiaoquJsonArray.Count; k++)
                            {
                                JObject xiaoquJson = xiaoquJsonArray[k] as JObject;
                                string xiaoquName = CommonUtil.HtmlDecode(xiaoquJson["name"].ToString());
                                string sale_numStr = xiaoquJson["sale_num"].ToString().Trim();
                                int saleNum = sale_numStr.Length == 0 ? 0 : int.Parse(sale_numStr);
                                string xiaoquUrl = CommonUtil.HtmlDecode(xiaoquJson["url"].ToString());
                                if (!urlDic.ContainsKey(xiaoquUrl) && saleNum > 0)
                                {
                                    urlDic.Add(xiaoquUrl, null);
                                    int pageIndex = 0;
                                    while (pageIndex * 60 < saleNum)
                                    {
                                        Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                        string ershoufangPageUrl = xiaoquUrl + "esf/?page=" + (pageIndex + 1).ToString();
                                        f2vs.Add("detailPageUrl", ershoufangPageUrl);
                                        f2vs.Add("detailPageName", ershoufangPageUrl);
                                        f2vs.Add("saleNum", sale_numStr);
                                        f2vs.Add("xiaoquName", xiaoquName);
                                        f2vs.Add("xiaoquUrl", xiaoquUrl);
                                        f2vs.Add("cityName", row["cityName"]);
                                        f2vs.Add("cityCode", row["cityCode"]);
                                        f2vs.Add("level1AreaName", row["level1AreaName"]);
                                        f2vs.Add("level1AreaCode", row["level1AreaCode"]);
                                        f2vs.Add("level2AreaCode", row["level2AreaCode"]);
                                        f2vs.Add("level2AreaName", row["level2AreaName"]);
                                        resultEW.AddRow(f2vs);
                                        pageIndex++;
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            throw ex;
                        }
                    }
                }
            }
            foreach (string cityName in cityToEW.Keys)
            {
                ExcelWriter resultEW = cityToEW[cityName];
                resultEW.SaveToDisk();
            }
        }
    }
}