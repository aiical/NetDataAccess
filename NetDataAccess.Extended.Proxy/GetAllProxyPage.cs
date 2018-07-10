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

namespace NetDataAccess.Extended.Proxy
{
    public class GetAllProxyPage : ExternalRunWebPage
    { 
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            try
            {
                string[] pSplits = this.Parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);

                CheckThreadCount = int.Parse(pSplits[0]);
                _Timeout = int.Parse(pSplits[1]);
                _TestPageUrl = pSplits[2];
                _CheckText = pSplits[3];
                _PageEncoding = pSplits[4];

                String exportDir = this.RunPage.GetExportDir();

                Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
                //resultColumnDic.Add("detailPageUrl", 0);
                //resultColumnDic.Add("detailPageName", 1);
                //resultColumnDic.Add("cookie", 2);
                //resultColumnDic.Add("grabStatus", 3);
                //resultColumnDic.Add("giveUpGrab", 4);
                resultColumnDic.Add("ip", 0);
                resultColumnDic.Add("port", 1);
                resultColumnDic.Add("user", 2);
                resultColumnDic.Add("pwd", 3);
                resultColumnDic.Add("usable", 4); 
                string resultFilePath = Path.Combine(exportDir, "Proxy.xlsx");

                ExcelWriter resultEW = new ExcelWriter(resultFilePath, "Proxy", resultColumnDic, null);

                string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;

                Dictionary<string, string> ipPortDic = new Dictionary<string, string>();
                List<Dictionary<string, string>> allIpPortList = new List<Dictionary<string, string>>();

                for (int i = 0; i < listSheet.RowCount; i++)
                {
                    Dictionary<string, string> row = listSheet.GetRow(i);
                    bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                    if (!giveUp)
                    {
                        string url = row[detailPageUrlColumnName];
                        string siteName = row["sitename"];
                        HtmlAgilityPack.HtmlDocument pageHtmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                        List<Dictionary<string, string>> ipPortList = getUsableProxies(url, siteName, pageHtmlDoc);
                        foreach (Dictionary<string, string> ipPort in ipPortList)
                        {
                            string key = ipPort["ip"] + ":" + ipPort["port"];
                            if (!ipPortDic.ContainsKey(key))
                            {
                                ipPortDic.Add(key, null);
                                AllIpPortList.Add(ipPort);
                            }
                        }
                    }
                }

                this.CheckAllIpPortUsable();
                while (FinishedCheckThreadCount < CheckThreadCount)
                {
                    Thread.Sleep(1000);
                }

                string msgInfo = "检测完毕，共获得可用代理服务器数:" + UsableIpPortList.Count.ToString();
                this.RunPage.InvokeAppendLogText(msgInfo, LogLevelType.Normal, true);

                foreach (Dictionary<string, string> ipPort in UsableIpPortList)
                {
                    ipPort.Add("usable", "是");
                    Dictionary<string, string> ipItem = new Dictionary<string, string>();
                    ipItem.Add("ip", ipPort["ip"]);
                    ipItem.Add("port", ipPort["port"]);
                    ipItem.Add("usable", "是"); 

                    resultEW.AddRow(ipItem);
                }

                resultEW.SaveToDisk();

                return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private List<Dictionary<string, string>> UsableIpPortList = new List<Dictionary<string, string>>();
        private List<Dictionary<string, string>> AllIpPortList = new List<Dictionary<string, string>>();

        private int _Timeout = 30000;
        private string _TestPageUrl = "";
        private string _CheckText = "";
        private string _PageEncoding = "";

        private int FinishedCheckThreadCount = 0;
        private int CheckThreadCount = 1;

        private object _AllListLocker = new object();
        private Dictionary<string, string> GetOneIpPort()
        {
            lock (_AllListLocker)
            {
                if (AllIpPortList.Count == 0)
                {
                    return null;
                }
                else
                {
                    string msgInfo = "待检测数:" + AllIpPortList.Count.ToString();
                    this.RunPage.InvokeAppendLogText(msgInfo, LogLevelType.Normal, true);

                    Dictionary<string, string> ipPort = AllIpPortList[0];
                    AllIpPortList.RemoveAt(0);
                    return ipPort;
                }
            }
        }

        private void CheckAllIpPortUsable()
        {
            for (int i = 0; i < CheckThreadCount; i++)
            {
                Thread checkThead = new Thread(new ThreadStart(CheckIpPortUsableThread));
                checkThead.Start();
                Thread.Sleep(100);
            }
        }

        private void CheckIpPortUsableThread()
        {
            while (1 == 1)
            {
                Dictionary<string, string> ipPort = this.GetOneIpPort();
                if (ipPort == null)
                {
                    FinishedCheckThreadCount++;
                    break;
                }
                else
                {
                    string ip = ipPort["ip"];
                    string port = ipPort["port"];
                    Nullable<TimeSpan> ts = CheckIpPortUsable(ip, port, _Timeout, _TestPageUrl, _CheckText);
                    this.RunPage.InvokeAppendLogText("监测获得可用代理数: " + UsableIpPortList.Count.ToString(), LogLevelType.System, true);
                    if(ts != null)
                    {
                        ipPort.Add("timespan", ((TimeSpan)ts).TotalMilliseconds.ToString());
                        ipPort.Add("detailPageUrl", "http://www.ip138.com/ips138.asp?ip=" + ip + "&port=" + port);
                        ipPort.Add("detailPageName", ip + ":" + port);
                        UsableIpPortList.Add(ipPort);
                    }
                }
            }
        }

        private Dictionary<string, string> _CheckResults = new Dictionary<string, string>();

        private Nullable<TimeSpan> CheckIpPortUsable(string ip, string port, int timeout, string testPageUrl,string checkText)
        {
            NDAWebClient client = null;
            try
            {
                DateTime dt1 = DateTime.Now;
                client = new NDAWebClient();
                client.Id = ip + ":" + port;
                System.Net.ServicePointManager.DefaultConnectionLimit = 512;
                client.Timeout = timeout;
                WebProxy webProxy = new WebProxy(ip, int.Parse(port));
                string ipPort  = webProxy.Address.Authority;
                client.Proxy = webProxy; 
                client.OpenReadCompleted += client_OpenReadCompleted;
                client.OpenReadAsync(new Uri(testPageUrl)); 
                string s = null;
                int waitingTime = 0;
                while (s == null && waitingTime < timeout)
                {
                    if (_CheckResults.ContainsKey(ipPort))
                    {
                        s = _CheckResults[ipPort];
                    }
                    else
                    {
                        waitingTime = waitingTime + 1000;
                        Thread.Sleep(1000);
                    }
                }

                if (s != null && s.Contains(checkText))
                {
                    DateTime dt2 = DateTime.Now;
                    string msgInfo = "检测成功!";
                    //this.InvokeAppendLogText(errorInfo, LogLevelType.Error, true);
                    this.RunPage.InvokeAppendLogText(msgInfo, LogLevelType.Normal, true);
                    return dt2 - dt1;
                }
                else
                {
                    DateTime dt2 = DateTime.Now;
                    string msgInfo = "检测超时!";
                    this.RunPage.InvokeAppendLogText(msgInfo, LogLevelType.Error, true);
                    return null;
                }
            }
            catch (Exception ex)
            {
                DateTime dt2 = DateTime.Now;
                string errorInfo = "访问失败, " + ip + ":" + port + ". " + ex.Message;
                //this.InvokeAppendLogText(errorInfo, LogLevelType.Error, true);
                this.RunPage.InvokeAppendLogText(errorInfo, LogLevelType.Error, true);
                return null;
            }
        }

        private void client_OpenReadCompleted(object sender, OpenReadCompletedEventArgs e)
        {
            Stream data = null;
            StreamReader reader = null;
            try
            {
                NDAWebClient client = (NDAWebClient)sender;
                string ipPort = ((WebProxy)client.Proxy).Address.Authority;
                reader = new StreamReader(e.Result, Encoding.GetEncoding(_PageEncoding));
                string s = reader.ReadToEnd();
                _CheckResults.Add(ipPort, s);
            }
            catch (Exception ex)
            {

            }
            finally
            {
                if (data != null)
                {
                    data.Close();
                    data.Dispose();
                }
                if (reader != null)
                {
                    reader.Close();
                    reader.Dispose();
                }
            }
        }

        private List<Dictionary<string, string>> getUsableProxies(string url, string siteName, HtmlAgilityPack.HtmlDocument pageHtmlDoc)
        {
            List<Dictionary<string, string>> ipPortList = new List<Dictionary<string, string>>();

            switch(siteName){
                case "chnlanker":
                    #region http://proxy.chnlanker.com/
                    {
                        HtmlNodeCollection trNodes = pageHtmlDoc.DocumentNode.SelectNodes("//div[@class=\"txt\"]/font/table/tr");
                        if (trNodes != null)
                        {
                            for (int i = 1; i < trNodes.Count; i++)
                            {
                                HtmlNode trNode = trNodes[i];
                                HtmlNodeCollection tdNodes = trNode.SelectNodes("./b/td");
                                string ip = tdNodes[1].InnerText.Trim();
                                string port = tdNodes[2].InnerText.Trim();
                                Dictionary<string, string> ipPort = new Dictionary<string, string>();
                                ipPort.Add("ip", ip);
                                ipPort.Add("port", port);
                                ipPort.Add("fromSiteName", siteName);
                                ipPortList.Add(ipPort);
                            }
                        }
                    }
                    #endregion
                    break;
                case "xicidaili": 
                    #region http://www.xicidaili.com/nn/1
                    {
                        HtmlNodeCollection trNodes = pageHtmlDoc.DocumentNode.SelectNodes("//table[@id=\"ip_list\"]/tr");
                        if (trNodes != null)
                        {
                            for (int i = 1; i < trNodes.Count; i++)
                            {
                                HtmlNode trNode = trNodes[i];
                                HtmlNodeCollection tdNodes = trNode.SelectNodes("./td");
                                string ip = tdNodes[1].InnerText.Trim();
                                string port = tdNodes[2].InnerText.Trim();
                                Dictionary<string, string> ipPort = new Dictionary<string, string>();
                                ipPort.Add("ip", ip);
                                ipPort.Add("port", port);
                                ipPort.Add("fromSiteName", siteName);
                                ipPortList.Add(ipPort);
                            }
                        }
                    }
                    #endregion
                    break;
                case "kuaidaili":
                    #region http://www.kuaidaili.com/free/inha/1
                    {
                        HtmlNodeCollection trNodes = pageHtmlDoc.DocumentNode.SelectNodes("//div[@id=\"list\"]/table[1]/tbody/tr");
                        if (trNodes != null)
                        {
                            for (int i = 1; i < trNodes.Count; i++)
                            {
                                HtmlNode trNode = trNodes[i];
                                HtmlNodeCollection tdNodes = trNode.SelectNodes("./td");
                                string ip = tdNodes[0].InnerText.Trim();
                                string port = tdNodes[1].InnerText.Trim();
                                Dictionary<string, string> ipPort = new Dictionary<string, string>();
                                ipPort.Add("ip", ip);
                                ipPort.Add("port", port);
                                ipPort.Add("fromSiteName", siteName);
                                ipPortList.Add(ipPort);
                            }
                        }
                    }
                    #endregion
                    break;
                case "cn-proxy":
                    #region http://cn-proxy.com
                    {
                        HtmlNodeCollection trNodes = pageHtmlDoc.DocumentNode.SelectNodes("//div[@class=\"table-container\"]/table/tbody/tr");
                        if (trNodes != null)
                        {
                            for (int i = 1; i < trNodes.Count; i++)
                            {
                                HtmlNode trNode = trNodes[i];
                                HtmlNodeCollection tdNodes = trNode.SelectNodes("./td");
                                string ip = tdNodes[0].InnerText.Trim();
                                string port = tdNodes[1].InnerText.Trim();
                                Dictionary<string, string> ipPort = new Dictionary<string, string>();
                                ipPort.Add("ip", ip);
                                ipPort.Add("port", port);
                                ipPort.Add("fromSiteName", siteName);
                                ipPortList.Add(ipPort);
                            }
                        }
                    }
                    #endregion
                    break; 
            }
            return ipPortList;
        }
    }
}