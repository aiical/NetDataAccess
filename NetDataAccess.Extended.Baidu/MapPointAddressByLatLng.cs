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

namespace NetDataAccess.Extended.Baidu
{
    [System.Runtime.InteropServices.ComVisibleAttribute(true)]
    public class MapPointAddressByLatLng : CustomProgramBase
    {
        #region _Succeed 
        private Dictionary<string, bool> _RangeSucceed = new Dictionary<string,bool>();
        private int _ThreadEndCount = 0;
        private int _GrabThreadCount = 1; 
        private string _PageUrl = "";
        private string _ExportDir = "";
        private int _GrabedPointNum = 0;
        private int _RangeGrabTimeout = 0;
        private DateTime _StartTime;
        #endregion

        #region Run
        public bool Run(string parameters, IListSheet listSheet)
        {

            string[] parameterArray = parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            _PageUrl = parameterArray[0];
            _ExportDir = parameterArray[1];
            _GrabThreadCount = int.Parse(parameterArray[2]);
            _RangeGrabTimeout = int.Parse(parameterArray[3]);
            _ListSheet = listSheet;

            _StartTime = DateTime.Now;

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
            int needGrabCount = listSheet.RowCount-_GrabedPointNum;
            if (needGrabCount == 0)
            {
                SaveAllPointsToFile();
                return true;
            }
            else
            {
                this.RunPage.InvokeAppendLogText("尚有" + needGrabCount.ToString() + "个未抓取", LogLevelType.Error, true);
                return false;
            }
        }
        #endregion

        private object getLocker = new object();
        private Dictionary<string, string> GetNextListRow()
        {
            lock (getLocker)
            {
                Dictionary<string, string> listValues = null;
                string fileName = "";
                while (listValues == null && _CurrentIndex < _ListSheet.RowCount)
                {
                    listValues = _ListSheet.GetRow(_CurrentIndex);
                    fileName = listValues["detailPageName"];

                    string filePath = this.GetResultPath(fileName);
                    if (File.Exists(filePath))
                    {
                        listValues = null;
                        _GrabedPointNum++;
                        _CurrentIndex++;
                    }
                }
                _CurrentIndex++;
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
            Dictionary<string, string> listValues = this.GetNextListRow();
            bool gotLastTime = false;
            while (listValues != null)
            {
                _RangeSucceed[tabName] = false;
                BeginSearch(tabName, listValues, gotLastTime);
                int waitingSeconds = 0;
                while (!_RangeSucceed[tabName] && waitingSeconds < _RangeGrabTimeout)
                {
                    waitingSeconds = waitingSeconds + _GrabPerCheckSecond * 1000;
                    Thread.Sleep(_GrabPerCheckSecond * 1000);
                }
                if (waitingSeconds >= _RangeGrabTimeout)
                {
                    //超时
                    string name = listValues["name"];
                    this.RunPage.InvokeAppendLogText("抓取超时, " + name, LogLevelType.Error, true);
                    gotLastTime = false;
                }
                else
                {
                    gotLastTime = true;
                }
                listValues = this.GetNextListRow();
            }
            _ThreadEndCount++;
        }
        private void BeginSearch(string tabName, Dictionary<string, string> listValues, bool gotLastTime)
        {
            //允许跳转到查询页面的次数，有时会出现跳转至登录页面的情况
            const int allowGoToQueryPageCount = 10;

            int goToQueryPageErrorCount = 0;

            string currentUrl = "";
            WebBrowser webBrowser = null;
            if (gotLastTime)
            {
                //加载网页
                webBrowser = (WebBrowser)this.RunPage.GetWebBrowserByName(tabName);
                currentUrl = webBrowser.Url.ToString();
            }
            else
            {
                while (currentUrl != _PageUrl && goToQueryPageErrorCount < allowGoToQueryPageCount)
                {
                    //加载网页
                    webBrowser = this.ShowWebPage(_PageUrl, tabName);
                    currentUrl = webBrowser.Url.ToString();
                }
            }

            if (currentUrl != _PageUrl)
            {
                throw new Exception("无法定位到查询页面.");
            }
            else
            {
                string uid = listValues["uid"];
                string lng = listValues["lng"];
                string lat = listValues["lat"];
                InvokeSearchMapPoint(webBrowser, lng, lat, uid, tabName);
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
            WebBrowser webBrowser = (WebBrowser)this.RunPage.InvokeShowWebPage(url, tabName, WebBrowserType.IE); 
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
        private string InvokeSearchMapPoint(WebBrowser webBrowser, string lng, string lat, string uid, string tabName)
        {
            string districts = (string)webBrowser.Invoke(new SearchMapPointInvokeDelegate(SearchMapPoint), new object[] { webBrowser, lng, lat, uid, tabName });
            return districts;
        }
        private delegate void SearchMapPointInvokeDelegate(WebBrowser webBrowser, string lng, string lat, string uid, string tabName);
        private void SearchMapPoint(WebBrowser webBrowser, string lng, string lat, string uid, string tabName)
        {
            webBrowser.Document.InvokeScript("searchMapPoint", new object[] { lng, lat, uid, tabName }); 
        }
        #endregion

        #region AfterGetPointResults 
        public void AfterGetPointResults(object item, string tabName, string fileName)
        { 
            IReflect itemReflect = item as IReflect;
            Dictionary<string, string> p2vs = new Dictionary<string, string>();
            p2vs.Add("district", (string)itemReflect.InvokeMember("district", BindingFlags.GetProperty, null, item, null, null, null, null));
            p2vs.Add("street", (string)itemReflect.InvokeMember("street", BindingFlags.GetProperty, null, item, null, null, null, null));
            p2vs.Add("province", (string)itemReflect.InvokeMember("province", BindingFlags.GetProperty, null, item, null, null, null, null));
            p2vs.Add("city", itemReflect.InvokeMember("city", BindingFlags.GetProperty, null, item, null, null, null, null).ToString());
            p2vs.Add("streetNumber", (string)itemReflect.InvokeMember("streetNumber", BindingFlags.GetProperty, null, item, null, null, null, null));
            SavePointToFile(p2vs, fileName);
            _GrabedPointNum++;
            _RangeSucceed[tabName] = true;
        }
        #endregion 

        #region SavePointsToFile
        private void SavePointToFile(Dictionary<string, string> point, string fileName)
        {
            string uid = fileName;
            JObject pointObj = new JObject();
            pointObj.Add("district", point["district"]);
            pointObj.Add("street", point["street"]);
            pointObj.Add("province", point["province"]);
            pointObj.Add("city", point["city"]);
            pointObj.Add("streetNumber", point["streetNumber"]);
            string filePath = this.GetResultPath(fileName);
            CommonUtil.CreateFileDirectory(filePath);
            FileHelper.SaveTextToFile(pointObj.ToString(), filePath);
        }

        private bool SaveAllPointsToFile()
        {  
            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                "uid",
                "title",
                "address",
                "province",
                "city",
                "district",
                "street", 
                "streetNumber",
                "phoneNumber",
                "postcode",
                "lat",
                "lng", 
                "detailUrl",
                "url"});
            string resultFilePath = Path.Combine(_ExportDir, "百度地图_包括行政区街道.xlsx"); 
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);

            SavePointsToFile(_ListSheet, resultEW);
            resultEW.SaveToDisk();

            return true;
        }

        private String GetResultPath(string name)
        {
            string resultFilePath = Path.Combine(_ExportDir, name + ".xlsx");
            return resultFilePath;
        }

        private void SavePointsToFile(IListSheet listSheet, ExcelWriter resultEW)
        {
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                //listSheet中只有一条记录
                string pageUrl = listSheet.PageUrlList[i];
                Dictionary<string, string> row = listSheet.GetRow(i);
                string uid = row["uid"];

                string localFilePath = this.GetResultPath(uid);
                string fileText = FileHelper.GetTextFromFile(localFilePath);
                int jsonBeginIndex = fileText.IndexOf("{");
                int jsonEndIndex = fileText.LastIndexOf("}");
                string jsonStr = fileText.Substring(jsonBeginIndex, jsonEndIndex - jsonBeginIndex + 1);
                JObject rootJo = JObject.Parse(jsonStr);
                string district = rootJo.GetValue("district") == null ? "" : rootJo.GetValue("district").ToString();
                string street = rootJo.GetValue("street") == null ? "" : rootJo.GetValue("street").ToString();
                string streetNumber = rootJo.GetValue("streetNumber") == null ? "" : rootJo.GetValue("streetNumber").ToString();

                Dictionary<string, string> p2vs = new Dictionary<string, string>();
                p2vs.Add("address", row["address"]);
                p2vs.Add("city", row["city"]);
                p2vs.Add("detailUrl", row["detailUrl"]); 
                p2vs.Add("phoneNumber", row["phoneNumber"]);
                p2vs.Add("postcode", row["postcode"]);
                p2vs.Add("province", row["province"]);
                p2vs.Add("title", row["title"]);
                p2vs.Add("uid", row["uid"]);
                p2vs.Add("url", row["url"]);
                p2vs.Add("lat", row["lat"]);
                p2vs.Add("lng", row["lng"]); 
                p2vs.Add("district", district);
                p2vs.Add("street", street);
                p2vs.Add("streetNumber", streetNumber);

                resultEW.AddRow(p2vs);
            }
        }
        #endregion
    }
}