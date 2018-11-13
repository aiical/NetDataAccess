using NetDataAccess.Base.Browser;
using NetDataAccess.Base.DB;
using NetDataAccess.Base.Definition;
using NetDataAccess.Base.Web;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NetDataAccess.Base.UI
{
    public interface IExternalRunWebPage
    {
        /// <summary>
        /// 在执行所有的爬取操作前
        /// </summary>
        /// <returns></returns>
        bool BeforeAllGrab();

        /// <summary>
        /// 初始化
        /// </summary>
        /// <param name="runPage"></param>
        /// <param name="parameters"></param>
        void Init(IRunWebPage runPage, string parameters);

        /// <summary>
        /// 在执行所有的爬取操作后
        /// </summary>
        /// <param name="listSheet"></param>
        /// <returns></returns>
        bool AfterAllGrab(IListSheet listSheet);

        /// <summary>
        /// 网页在浏览器中加载成功后（通过浏览器获取页面）
        /// </summary>
        /// <param name="pageUrl"></param>
        /// <param name="listRow"></param>
        /// <param name="webBrowser"></param>
        void WebBrowserHtml_AfterPageLoaded(string pageUrl, Dictionary<string, string> listRow, IWebBrowser webBrowser);

        /// <summary>
        /// 发送请求前（通过httprequest获取页面）
        /// </summary>
        /// <param name="pageUrl"></param>
        /// <param name="listRow"></param>
        /// <param name="client"></param>
        void WebRequestHtml_BeforeSendRequest(string pageUrl, Dictionary<string, string> listRow, NDAWebClient client);

        /// <summary>
        /// 当抓取了一个页面后
        /// </summary>
        /// <param name="pageUrl"></param>
        /// <param name="listRow"></param>
        /// <param name="needReGrab">是否需要重新抓取，当抓取成功或者放弃抓取时，此值为false</param>
        /// <param name="existLocalFile">抓取前已经存在了本地文件</param> 
        void AfterGrabOne(string pageUrl, Dictionary<string, string> listRow, bool needReGrab, bool existLocalFile);

        /// <summary>
        /// 当抓取了一个页面出错后
        /// </summary>
        /// <param name="pageUrl"></param>
        /// <param name="listRow"></param>
        /// <param name="ex"></param>
        /// <returns>是否放弃抓取</returns>
        bool AfterGrabOneCatchException(string pageUrl, Dictionary<string, string> listRow, Exception ex);

        /// <summary>
        /// 当抓取了一个页面前
        /// </summary>
        /// <param name="pageUrl"></param>
        /// <param name="listRow"></param> 
        /// <param name="existLocalFile">抓取前已经存在了本地文件</param> 
        void BeforeGrabOne(string pageUrl, Dictionary<string, string> listRow, bool existLocalFile);

        /// <summary>
        /// 通过其它方式获取数据
        /// </summary>
        /// <param name="listRow"></param>
        void GetDataByOtherAccessType(Dictionary<string, string> listRow);

        /// <summary>
        /// 是否需要抓取
        /// </summary>
        /// <param name="listRow"></param>
        /// <param name="localPagePath"></param>
        /// <returns></returns>
        bool CheckNeedGrab(Dictionary<string, string> listRow, string localPagePath);

        /// <summary>
        /// 获取要传递给服务器端的信息
        /// </summary>
        /// <param name="pageUrl"></param>
        /// <param name="listRow"></param>
        /// <param name="encoding"></param>
        /// <returns></returns>
        byte[] GetRequestData_BeforeSendRequest(string pageUrl, Dictionary<string, string> listRow, Encoding encoding);

        /// <summary>
        /// 判断是否抓取单个页面完成
        /// </summary>
        /// <param name="webPageText"></param>
        /// <param name="listRow"></param>
        void CheckRequestCompleteFile(string webPageText, Dictionary<string, string> listRow);

        /// <summary>
        /// 自定义ProgramType的方式，逐个抓取详情页
        /// </summary>
        /// <param name="listSheet"></param>
        /// <param name="detailPageInfo"></param>
        /// <returns></returns>
        bool BeginGrabDetailPageInExternalProgram(IListSheet listSheet, Proj_Detail_SingleLine detailPageInfo);


        void WebBrowserHtml_AfterDoNavigate(string pageUrl, Dictionary<string, string> listRow, string tabName);
    }
    
}
