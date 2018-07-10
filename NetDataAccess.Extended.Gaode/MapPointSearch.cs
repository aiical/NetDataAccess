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

namespace NetDataAccess.Extended.Gaode
{
    [System.Runtime.InteropServices.ComVisibleAttribute(true)]
    public class MapPointSearch : ExternalRunWebPage
    {
        #region _Succeed 
        private Dictionary<string, bool> _RangeSucceed = new Dictionary<string,bool>();
        private int _ThreadEndCount = 0;
        private int _ThreadErrorCount = 0;
        private int _GrabThreadCount = 1;
        private string _PageUrl = "";
        private string _ExportDir = "";
        private int _GrabPointNum = 0;
        private int _RangeGrabTimeout = 0;
        private DateTime _StartTime;
        private int OnePageRowCount = 100;
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
        private List<Dictionary<string, string>> GetNextPackageListRow()
        {
            lock (getLocker)
            {
                string fileName = "";
                List<Dictionary<string, string>> listValues = null;
                while (listValues == null && _CurrentIndex < _ListSheet.RowCount)
                {
                    Dictionary<string, string> listValue = _ListSheet.GetRow(_CurrentIndex);
                    fileName = listValue["detailPageName"];

                    string filePath = this.GetResultPath(fileName, "csv");
                    if (File.Exists(filePath))
                    {
                        listValues = null;
                        _CurrentIndex = _CurrentIndex + OnePageRowCount;
                    }
                    else
                    {
                        listValues = new List<Dictionary<string, string>>();
                        for (int i = _CurrentIndex; i < _CurrentIndex + OnePageRowCount; i++)
                        {
                            if (i < _ListSheet.RowCount)
                            {
                                listValues.Add(_ListSheet.GetRow(i));
                            }
                        }
                    }
                }
                _CurrentIndex = _CurrentIndex + OnePageRowCount;
                return listValues;
            }

        }

        private IListSheet _ListSheet = null;
        private int _CurrentIndex = 0;
        private int _GrabPerCheckSecond = 2;

        #region 开始搜索
        private void BeginSearch(object tabNameObj)
        {
            string tabName = (string)tabNameObj;
            _RangeSucceed[tabName] = false;
            List<Dictionary<string, string>> listValues = this.GetNextPackageListRow();
            while (listValues != null)
            {
                try
                {
                    _RangeSucceed[tabName] = false;
                    BeginSearch(tabName, listValues);
                    int waitingSeconds = 0;
                    while (!_RangeSucceed[tabName] && waitingSeconds < _RangeGrabTimeout)
                    {
                        waitingSeconds = waitingSeconds + _GrabPerCheckSecond;
                        Thread.Sleep(_GrabPerCheckSecond * 1000);
                    }
                    if (waitingSeconds >= _RangeGrabTimeout)
                    {
                        //超时
                        string name = listValues[0]["name"];
                        this.RunPage.InvokeAppendLogText("抓取超时, " + name, LogLevelType.Error, true);
                        _ThreadErrorCount++;
                    }
                }
                catch (Exception ex)
                {
                    //超时
                    string name = listValues[0]["name"];
                    this.RunPage.InvokeAppendLogText(ex.Message + ". name", LogLevelType.Error, true);
                    _ThreadErrorCount++;
                }
                listValues = this.GetNextPackageListRow();
            }
            _ThreadEndCount++;
        }
        private void BeginSearch(string tabName, List<Dictionary<string, string>> listValues)
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

            if (currentUrl.ToLower() != _PageUrl.ToLower())
            {
                throw new Exception("无法定位到查询页面.");
            }
            else
            {
                string fileName = listValues[0]["detailPageName"];
                StringBuilder allPointStrBuilder = new StringBuilder();
                for (int i = 0; i < listValues.Count; i++)
                {
                    Dictionary<string, string> listValue = listValues[i];
                    string lat = listValue["lat"];
                    string lng = listValue["lng"];
                    allPointStrBuilder.Append(lng + "," + lat + ";");
                }

                this.RunPage.InvokeAppendLogText("（" + _CurrentIndex.ToString() + "/" + _ListSheet.RowCount.ToString() + "）准备抓取区域：" + fileName + " (包括" + OnePageRowCount.ToString() + "个点)", LogLevelType.Normal, true);
                InvokeSearchMap(webBrowser, allPointStrBuilder.ToString(), fileName, tabName);
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
        private int WebRequestTimeout = 20 * 1000;
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
        private string InvokeSearchMap(WebBrowser webBrowser, string allPointStr, string fileName, string tabName)
        {
            string districts = (string)webBrowser.Invoke(new SearchMapInvokeDelegate(SearchMap), new object[] { webBrowser, allPointStr, fileName, tabName });
            return districts;
        }
        private delegate void SearchMapInvokeDelegate(WebBrowser webBrowser, string allPointStr, string fileName, string tabName);
        private void SearchMap(WebBrowser webBrowser, string allPointStr, string fileName, string tabName)
        {
            webBrowser.Document.InvokeScript("searchMap", new object[] { allPointStr, fileName, tabName }); 
        }
        #endregion 

        #region AfterGetPoints
        public void AfterGetPoints(object obj, string fileName, string tabName)
        {
            this.RunPage.InvokeAppendLogText("该批次抓取完成", LogLevelType.System, true);

            IReflect arrReflect = obj as IReflect;
            int length = (int)arrReflect.InvokeMember("length", BindingFlags.GetProperty, null, obj, null, null, null, null);
            List<Dictionary<string, string>> pointInfos = new List<Dictionary<string, string>>();
            for (int i = 0; i < length; i++)
            {
                object item = (object)arrReflect.InvokeMember(i.ToString(), BindingFlags.GetProperty, null, obj, null, null, null, null);
                IReflect itemReflect = item as IReflect;
                Dictionary<string, string> p2vs = new Dictionary<string, string>();
                p2vs.Add("city", (string)itemReflect.InvokeMember("city", BindingFlags.GetProperty, null, item, null, null, null, null));
                p2vs.Add("district", (string)itemReflect.InvokeMember("district", BindingFlags.GetProperty, null, item, null, null, null, null));
                p2vs.Add("province", (string)itemReflect.InvokeMember("province", BindingFlags.GetProperty, null, item, null, null, null, null));
                p2vs.Add("street", (string)itemReflect.InvokeMember("street", BindingFlags.GetProperty, null, item, null, null, null, null));
                p2vs.Add("streetNumber", itemReflect.InvokeMember("streetNumber", BindingFlags.GetProperty, null, item, null, null, null, null).ToString());
                p2vs.Add("township", itemReflect.InvokeMember("township", BindingFlags.GetProperty, null, item, null, null, null, null).ToString());
                p2vs.Add("formattedAddress", itemReflect.InvokeMember("formattedAddress", BindingFlags.GetProperty, null, item, null, null, null, null).ToString());
                p2vs.Add("building", itemReflect.InvokeMember("building", BindingFlags.GetProperty, null, item, null, null, null, null).ToString());
                p2vs.Add("buildingType", itemReflect.InvokeMember("buildingType", BindingFlags.GetProperty, null, item, null, null, null, null).ToString());
                p2vs.Add("lat", itemReflect.InvokeMember("lat", BindingFlags.GetProperty, null, item, null, null, null, null).ToString());
                p2vs.Add("lng", itemReflect.InvokeMember("lng", BindingFlags.GetProperty, null, item, null, null, null, null).ToString());

                pointInfos.Add(p2vs);
            }

            if (SaveRangePointsToFile(pointInfos, fileName))
            {
                _GrabPointNum = _GrabPointNum + OnePageRowCount;
                _RangeSucceed[tabName] = true;

                TimeSpan ts = DateTime.Now - _StartTime;
                if (_ListSheet.RowCount > _CurrentIndex)
                {
                    double needMinutes = ((double)_ListSheet.RowCount - (double)_CurrentIndex) * (ts.TotalMinutes / (double)_GrabPointNum);
                    this.RunPage.InvokeAppendLogText("预计剩余" + ((int)needMinutes).ToString() + "分钟", Base.EnumTypes.LogLevelType.System, true);
                }
            }
        }
        #endregion

        #region SavePointsToFile
        private bool SaveRangePointsToFile(List<Dictionary<string, string>> points, string fileName)
        {
            bool succeed = true;
            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                "province",
                "city",
                "district",
                "street",
                "streetNumber",
                "township",
                "formattedAddress",
                "building",
                "buildingType",
                "lat",
                "lng"});
            string resultFilePath = this.GetResultPath(fileName, "csv");
            CsvWriter resultEW = new CsvWriter(resultFilePath, resultColumnDic);

            SavePointsToFile(points, resultEW);

            return succeed;
        }

        private bool SaveAllPointsToFile()
        {
            List<Dictionary<string, string>> allPoints = new List<Dictionary<string, string>>();
            Dictionary<string, string> uidDic = new Dictionary<string, string>();
            for (int i = 0; i < _ListSheet.RowCount; i++)
            {
                if (i % OnePageRowCount == 0)
                {
                    Dictionary<string, string> listValues = _ListSheet.GetRow(i);
                    string name = listValues["detailPageName"];
                    string rangeResultFilePath = this.GetResultPath(name, "csv");

                    CsvReader er = new CsvReader(rangeResultFilePath);

                    for (int j = 0; j < er.GetRowCount(); j++)
                    {
                        Dictionary<string, string> row = er.GetFieldValues(j);
                        Dictionary<string, string> p2vs = new Dictionary<string, string>();
                        p2vs.Add("city", row["city"]);
                        p2vs.Add("district", row["district"]);
                        p2vs.Add("province", row["province"]);
                        p2vs.Add("street", row["street"]);
                        p2vs.Add("streetNumber", row["streetNumber"]);
                        p2vs.Add("township", row["township"]);
                        p2vs.Add("formattedAddress", row["formattedAddress"]);
                        p2vs.Add("building", row["building"]);
                        p2vs.Add("buildingType", row["buildingType"]);
                        p2vs.Add("lat", row["lat"]);
                        p2vs.Add("lng", row["lng"]);

                        allPoints.Add(p2vs);
                    }
                }
            }

            bool succeed = true;
            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{ 
                "province",
                "city",
                "district",
                "street",
                "streetNumber",
                "township",
                "formattedAddress",
                "building",
                "buildingType",
                "lat",
                "lng"});
            string resultFilePath = this.GetResultPath("百度地图爬取结果", "csv");
            CsvWriter resultEW = new CsvWriter(resultFilePath, resultColumnDic);

            SavePointsToFile(allPoints, resultEW);

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