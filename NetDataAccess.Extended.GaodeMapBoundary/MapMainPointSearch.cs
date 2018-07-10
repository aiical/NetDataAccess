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
using NetDataAccess.Base.Writer;
using NetDataAccess.Base.Reader;
using NetDataAccess.Base.CsvHelper;
using NetDataAccess.Base.DB;
using HtmlAgilityPack;
using System.Web;
using System.Runtime.Remoting;
using System.Reflection;
using System.Collections;

namespace NetDataAccess.Extended.GaodeMapBoundary
{
    [System.Runtime.InteropServices.ComVisibleAttribute(true)]
    public class MapMainPointSearch : ExternalRunWebPage
    {
        #region _Succeed 
        private Dictionary<string, bool> _RangeSucceed = new Dictionary<string,bool>();
        private int _ThreadEndCount = 0;
        private int _ThreadErrorCount = 0;
        private int _GrabThreadCount = 1;
        private string _PageUrl = "";
        private string _ExportDir = "";
        private int _GrabRangeNum = 0;
        private int _RangeGrabTimeout = 0;
        private DateTime _StartTime;
        #endregion

        #region Run
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            return this.BeginGrabAll(listSheet);
        }
        private bool BeginGrabAll(IListSheet listSheet)
        {
            string[] parameterArray = this.Parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            _PageUrl = parameterArray[0];
            _ExportDir = parameterArray[1];
            _GrabThreadCount = int.Parse(parameterArray[2]);
            _RangeGrabTimeout = int.Parse(parameterArray[3]);
            _ListSheet = listSheet;

            _StartTime = DateTime.Now;
            _CurrentIndex = 0;
            _ThreadErrorCount = 0;

            for (int i = 0; i < _GrabThreadCount; i++)
            {
                Thread grabThread = new Thread(new ParameterizedThreadStart(BeginSearch));
                grabThread.Start(i.ToString());
                Thread.Sleep(5000);
            }

            while (_ThreadEndCount < _GrabThreadCount)
            {
                Thread.Sleep(3000); 
            }

            if (_ThreadErrorCount > 0)
            {
                return this.BeginGrabAll(listSheet);
            }
            else
            {
                SaveAllPointsToFile();
            }
            return true;
        }
        #endregion

        private object getLocker = new object();
        private Dictionary<string, string> GetNextListRow()
        {
            lock (getLocker)
            {
                Dictionary<string, string> listValues = null;
                string fullName = "";
                while (listValues == null && _CurrentIndex < _ListSheet.RowCount)
                {
                    listValues = _ListSheet.GetRow(_CurrentIndex);
                    fullName = listValues["fullName"];

                    string filePath = this.GetResultPath(fullName, "txt");
                    if (File.Exists(filePath))
                    {
                        listValues = null;
                        _CurrentIndex++;
                    }
                }
                _CurrentIndex++;
                return listValues;
            }

        }

        private IListSheet _ListSheet = null;
        private int _CurrentIndex = 0;
        private int _GrabPerCheckSecond = 1000;

        #region 开始搜索
        private void BeginSearch(object tabNameObj)
        {
            string tabName = (string)tabNameObj;
            _RangeSucceed[tabName] = false;
            Dictionary<string, string> listValues = this.GetNextListRow();
            while (listValues != null)
            {
                _RangeSucceed[tabName] = false;
                BeginSearch(tabName, listValues);
                int waitingSeconds = 0;
                while (!_RangeSucceed[tabName] && waitingSeconds < _RangeGrabTimeout)
                {
                    waitingSeconds = waitingSeconds + _GrabPerCheckSecond;
                    Thread.Sleep(_GrabPerCheckSecond);
                }
                if (waitingSeconds >= _RangeGrabTimeout)
                {
                    //超时
                    string name = listValues["name"];
                    this.RunPage.InvokeAppendLogText("抓取超时, " + name, LogLevelType.Error, true);
                    _ThreadErrorCount++;
                }
                listValues = this.GetNextListRow();
            }
            _ThreadEndCount++;
        }
        private void BeginSearch(string tabName, Dictionary<string, string> listValues)
        {

            //允许跳转到查询页面的次数，有时会出现跳转至登录页面的情况
            const int allowGoToQueryPageCount = 10;

            int goToQueryPageErrorCount = 0;

            string currentUrl = "";
            WebBrowser webBrowser = null;

            while (currentUrl != _PageUrl && goToQueryPageErrorCount < allowGoToQueryPageCount)
            {
                //加载网页
                webBrowser = this.ShowWebPage(_PageUrl, tabName); 

                currentUrl = webBrowser.Url.ToString();
            }

            if (currentUrl != _PageUrl)
            {
                throw new Exception("无法定位到查询页面.");
            }
            else
            {
                string trimCode = listValues["trimCode"];
                string fullName = listValues["fullName"];
                string itemIndex = listValues["itemIndex"];

                this.RunPage.InvokeAppendLogText("（" + itemIndex + "/" + _ListSheet.RowCount.ToString() + "）准备抓取区域：" + trimCode + " " + fullName, LogLevelType.Normal, true);
                InvokeSearchMap(webBrowser, fullName, tabName);
            }
        }
        #endregion

        #region 浏览器控件
        /// <summary>
        /// 浏览器控件
        /// </summary>
        //private WebBrowser WebBrowserMain = null;
        #endregion 

        #region 获取网页信息超时时间
        /// <summary>
        /// 获取网页信息超时时间
        /// </summary>
        private int WebRequestTimeout = 60 * 1000;
        #endregion
        
        #region 显示网页
        private WebBrowser ShowWebPage(string url, string tabName)
        {
            WebBrowser webBrowser = this.RunPage.InvokeShowWebPage(url, tabName); 
            int waitCount = 0;
            while (!this.RunPage.CheckIsComplete(tabName))
            {
                if (SysConfig.WebPageRequestInterval * waitCount > WebRequestTimeout)
                {
                    string errorInfo = "打开页面超时! PageUrl = " + url + ". 但是继续执行!";
                    this.RunPage.InvokeAppendLogText(errorInfo, Base.EnumTypes.LogLevelType.System, true);
                    break;
                    //超时
                    //throw new Exception("打开页面超时. PageUrl = " + url);
                }
                //等待
                waitCount++;
                Thread.Sleep(SysConfig.WebPageRequestInterval);
            }


            this.InvokeAddMyScript(webBrowser);

            //再增加个等待，等待异步加载的数据
            Thread.Sleep(1000);
            return webBrowser;
        }
        #endregion

        #region AddMyScript
        private void InvokeAddMyScript(WebBrowser webBrowser)
        {
            webBrowser.Invoke(new AddMyScriptInvokeDelegate(AddMyScript), new object[] { webBrowser, "" });
        }
        private delegate void AddMyScriptInvokeDelegate(WebBrowser webBrowser, string p1);
        private void AddMyScript(WebBrowser webBrowser, string p1)
        {
            webBrowser.ObjectForScripting = this;
        }
        #endregion 

        #region 开始搜索
        private string InvokeSearchMap(WebBrowser webBrowser, string fullName, string tabName)
        {
            string districts = (string)webBrowser.Invoke(new SearchMapInvokeDelegate(SearchMap), new object[] { webBrowser, fullName, tabName });
            return districts;
        }
        private delegate void SearchMapInvokeDelegate(WebBrowser webBrowser, string fullName, string tabName);
        private void SearchMap(WebBrowser webBrowser, string fullName, string tabName)
        {
            webBrowser.Document.InvokeScript("searchMap", new object[] { fullName, tabName }); 
        }
        #endregion 

        #region AfterGetRange
        public void AfterGetPoint(string points, string fullName, string tabName)
        {
            if (SavePointsToFile(points, fullName))
            {
                _GrabRangeNum++;
                _RangeSucceed[tabName] = true;

                TimeSpan ts = DateTime.Now - _StartTime;
                if (_ListSheet.RowCount > _CurrentIndex)
                {
                    double needMinutes = ((double)_ListSheet.RowCount - (double)_CurrentIndex) * (ts.TotalMinutes / (double)_GrabRangeNum);
                    this.RunPage.InvokeAppendLogText("预计剩余" + ((int)needMinutes).ToString() + "分钟", Base.EnumTypes.LogLevelType.System, true);
                }
            }
        }
        #endregion

        #region SavePointsToFile
        private bool SavePointsToFile(string points, string name)
        {
            bool succeed = true; 
            string resultFilePath = this.GetResultPath(name, "txt");
            FileHelper.SaveTextToFile(points, resultFilePath);

            return succeed;
        }

        private bool SaveAllPointsToFile()
        {
            List<Dictionary<string, string>> allPoints = new List<Dictionary<string, string>>();

            StringBuilder str = new StringBuilder("var allDistrictMainPoints = {};");
            Dictionary<string, string> uidDic = new Dictionary<string, string>();
            for (int i = 0; i < _ListSheet.RowCount; i++)
            {
                Dictionary<string, string> listValues = _ListSheet.GetRow(i);
                string code = listValues["code"];
                string trimCode = listValues["trimCode"];
                string name = listValues["name"];
                string fullName = listValues["fullName"];
                string rangeResultFilePath = this.GetResultPath(fullName, "txt");

                string point = FileHelper.GetTextFromFile(rangeResultFilePath);

                Dictionary<string, string> p2vs = new Dictionary<string, string>();
                p2vs.Add("code", code);
                p2vs.Add("trimCode", trimCode);
                p2vs.Add("name", name);
                p2vs.Add("fullName", fullName);
                p2vs.Add("point", point);

                allPoints.Add(p2vs);

                string[] p = point.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                str.AppendLine("allDistrictMainPoints[\"" + code + "\"] = {code:\"" + code + "\", name: \"" + name + "\", point_x: " + p[0] + ", point_y: " + p[1] + "};");
            }

            bool succeed = true;
            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{ 
                "code", 
                "trimCode",
                "name",
                "fullName",
                "point"});
            string resultFilePath = this.GetResultPath("百度地图行政区划点", "csv");
            CsvWriter resultEW = new CsvWriter(resultFilePath, resultColumnDic);

            SavePointsToFile(allPoints, resultEW);
             
            string jsFilePath = this.GetResultPath("allDistrictMainPoints", "js");
            FileHelper.SaveTextToFile(str.ToString(), jsFilePath); 

            return succeed;
        } 

        private String GetResultPath(string name, string ext)
        {
            string resultFilePath = Path.Combine(_ExportDir, name + "." + ext);
            return resultFilePath;
        }

        private void SavePointsToFile(List<Dictionary<string, string>> points, CsvWriter resultEW)
        {
            for (int i = 0; i < points.Count; i++)
            {
                Dictionary<string, string> f2vs = points[i];
                resultEW.AddRow(f2vs);
            }
            resultEW.SaveToDisk();
        }

        private void SavePointsToFile(List<Dictionary<string, string>> points, ExcelWriter resultEW)
        {
            for (int i = 0; i < points.Count; i++)
            {
                Dictionary<string, string> f2vs = points[i];
                resultEW.AddRow(f2vs);
            }
            resultEW.SaveToDisk();
        }
        #endregion
    }
}