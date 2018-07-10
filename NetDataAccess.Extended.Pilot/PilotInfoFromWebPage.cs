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

namespace NetDataAccess.Extended.Pilot
{
    [System.Runtime.InteropServices.ComVisibleAttribute(true)]
    public class PilotInfoFromWebPage : ExternalRunWebPage
    {
        private Dictionary<string, string> _StateDic = null;
        private Dictionary<string, string> StateDic
        {
            get
            {
                if (_StateDic == null)
                {
                    Dictionary<string, string> stateDic = new Dictionary<string, string>();
                    stateDic.Add("AL", "Alabama");
                    stateDic.Add("AK", "Alaska");
                    stateDic.Add("AZ", "Arizona");
                    stateDic.Add("AR", "Arkansas");
                    stateDic.Add("CA", "California");
                    stateDic.Add("CO", "Colorado");
                    stateDic.Add("CT", "Connecticut");
                    stateDic.Add("DE", "Delaware");
                    stateDic.Add("FL", "Florida");
                    stateDic.Add("GA", "Georgia");
                    stateDic.Add("HI", "Hawaii");
                    stateDic.Add("ID", "Idaho");
                    stateDic.Add("IL", "Illinois");
                    stateDic.Add("IN", "Indiana");
                    stateDic.Add("IA", "Iowa");
                    stateDic.Add("KS", "Kansas");
                    stateDic.Add("KY", "Kentucky");
                    stateDic.Add("LA", "Louisiana");
                    stateDic.Add("ME", "Maine");
                    stateDic.Add("MD", "Maryland");
                    stateDic.Add("MA", "Massachusetts");
                    stateDic.Add("MI", "Michigan");
                    stateDic.Add("MN", "Minnesota");
                    stateDic.Add("MS", "Mississippi");
                    stateDic.Add("MO", "Missouri");
                    stateDic.Add("MT", "Montana");
                    stateDic.Add("NE", "Nebraska");
                    stateDic.Add("NV", "Nevada");
                    stateDic.Add("NH", "New hampshire");
                    stateDic.Add("NJ", "New jersey");
                    stateDic.Add("NM", "New mexico");
                    stateDic.Add("NY", "New York");
                    stateDic.Add("NC", "North Carolina");
                    stateDic.Add("ND", "North Dakota");
                    stateDic.Add("OH", "Ohio");
                    stateDic.Add("OK", "Oklahoma");
                    stateDic.Add("OR", "Oregon");
                    stateDic.Add("PA", "Pennsylvania");
                    stateDic.Add("RI", "Rhode island");
                    stateDic.Add("SC", "South carolina");
                    stateDic.Add("SD", "South dakota");
                    stateDic.Add("TN", "Tennessee");
                    stateDic.Add("TX", "Texas");
                    stateDic.Add("UT", "Utah");
                    stateDic.Add("VT", "Vermont");
                    stateDic.Add("VA", "Virginia");
                    stateDic.Add("WA", "Washington");
                    stateDic.Add("WV", "West Virginia");
                    stateDic.Add("WI", "Wisconsin");
                    stateDic.Add("WY", "Wyoming");
                    _StateDic = stateDic;
                }
                return _StateDic;
            }
        }

        #region Run
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            this.GetPilotInfo(listSheet);
            return true;
        }
        private void GetPilotInfo(IListSheet listSheet)
        { 
            
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("p_first_name", 0);
            resultColumnDic.Add("p_last_name", 1); 
            resultColumnDic.Add("certificate", 2);
            resultColumnDic.Add("date_of_issue", 3);
            string resultFilePath = Path.Combine(exportDir, "飞行员认证信息详情.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            Dictionary<string, string> urlDic = new Dictionary<string, string>();
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"];
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                { 
                    string localFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir);

                    try
                    {
                        HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                        htmlDoc.Load(localFilePath, Encoding.GetEncoding("utf-8"));
                        
                        String firstName = row["p_first_name"].Trim();
                        String lastName = row["p_last_name"].Trim();
                        string fullName = CommonUtil.StringArrayToString((firstName + " " + lastName).Split(new string[] { }, StringSplitOptions.RemoveEmptyEntries), " ");

                        int fullMatchNameCount = 0;
                        HtmlNodeCollection linkNodes = htmlDoc.DocumentNode.SelectNodes("//a[starts-with(@id, 'ctl00_content_ctl01_drAirmenList_ctl')]");
                        foreach (HtmlNode linkNode in linkNodes)
                        {
                            string[] nameParts = CommonUtil.HtmlDecode(linkNode.InnerText.Trim()).Trim().Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);
                            for (int j = 0; j < nameParts.Length; j++)
                            {
                                nameParts[j] = nameParts[j].Trim();
                            }
                            string fullNameInPage = CommonUtil.StringArrayToString(nameParts, " ");
                            if (fullNameInPage == fullName)
                            {
                                fullMatchNameCount++;
                            }
                        }

                        if (fullMatchNameCount == 1)
                        {
                            HtmlNodeCollection infoNodes = htmlDoc.DocumentNode.SelectNodes("//div[starts-with(@id, 'TabBody')]/label[@class='Cert_Info']");
                            for (int j = 0; j < infoNodes.Count; j++)
                            {
                                HtmlNode infoNode = infoNodes[j];
                                string text = CommonUtil.HtmlDecode(infoNode.InnerText.Trim());
                                if (text.Contains("Certificate:") && text.Contains("Date of Issue:"))
                                {
                                    int certStartIndex = text.IndexOf("Certificate:") + 12;
                                    int certEndIndex = text.IndexOf("Date of Issue:");
                                    int dateStartIndex = certEndIndex + 14;
                                    string certificate = text.Substring(certStartIndex, certEndIndex - certStartIndex).Trim();
                                    string date_of_issue = text.Substring(dateStartIndex).Trim();

                                    Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                    f2vs.Add("p_first_name", firstName);
                                    f2vs.Add("p_last_name", lastName);
                                    f2vs.Add("certificate", certificate);
                                    f2vs.Add("date_of_issue", date_of_issue);
                                    resultEW.AddRow(f2vs);
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
            resultEW.SaveToDisk();

        }
        #endregion

        private Dictionary<string, bool> _WebPageSucceeds = new Dictionary<string, bool>();

        public override void GetDataByOtherAcessType(Dictionary<string, string> listRow)
        {
            try
            {
                string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

                base.GetDataByOtherAcessType(listRow);
                String pageUrl = listRow[SysConfig.DetailPageUrlFieldName];
                String firstName = listRow["p_first_name"].Trim();
                String lastName = listRow["p_last_name"].Trim();
                string fullName = firstName + " " + lastName;
                int blankIndex = lastName.IndexOf(" ");
                if (blankIndex > 0)
                {
                    //如果lastname以JR、III等结尾，搜索不到，所以去掉
                    lastName = lastName.Substring(0, blankIndex);
                }  
                String state = listRow["p_state"].Trim();
                String city = listRow["p_city"].Trim();  

                String localFilePath = this.RunPage.GetFilePath(pageUrl, pageSourceDir);
                string threadId = Thread.CurrentThread.ManagedThreadId.ToString();
                _WebPageSucceeds[threadId] = false;
                BeginGetPilotInfo(threadId, pageUrl, firstName, lastName, "", "", "", "", fullName, state, "", localFilePath);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        #region 开始搜索
        private void BeginGetPilotInfo(string tabName, string pageUrl, string firstName, string lastName, string street1, string country, string med_date, string med_class, string fullName, string state, string city, string localFilePath)
        {
            string currentUrl = "";
            WebBrowser webBrowser = null;

            while (currentUrl != pageUrl)
            {
                //加载网页
                webBrowser = this.ShowWebPage(pageUrl, tabName);

                currentUrl = webBrowser.Url.ToString();
            }

            if (currentUrl != pageUrl)
            {
                throw new Exception("无法打开页面.");
            }

            this.InvokeInputNameScript(webBrowser, firstName, lastName, country, state);
            this.WaitGetPilotInfoPage(webBrowser);

            string pageHtml = this.RunPage.InvokeGetPageHtml(tabName);
            HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
            htmlDoc.LoadHtml(pageHtml);
            HtmlNodeCollection linkNodes = htmlDoc.DocumentNode.SelectNodes("//a[starts-with(@id, 'ctl00_content_ctl01_drAirmenList_ctl')]");
            List<String> linkNodeIds = new List<string>();
            if (linkNodes == null)
            {
                throw new CannotFoundException("没有找到对应人员");
                //throw new Exception("没有找到对应人员");
            }
            else
            {
                foreach (HtmlNode linkNode in linkNodes)
                {
                    string linkNodeId = linkNode.GetAttributeValue("id", "");
                    linkNodeIds.Add(linkNodeId);
                }

            }

            bool found = false;
            foreach (string linkNodeId in linkNodeIds)
            {
                this.InvokeClickPilotNameScript(webBrowser, linkNodeId);
                this.WaitGetPilotInfoPage(webBrowser);
                string pilotInfoPageHtml = this.RunPage.InvokeGetPageHtml(tabName);
                HtmlAgilityPack.HtmlDocument newHtmlDoc = new HtmlAgilityPack.HtmlDocument();
                newHtmlDoc.LoadHtml(pilotInfoPageHtml);
                HtmlNode cert_NameNode = newHtmlDoc.DocumentNode.SelectSingleNode("//div[@id='divPersonalInfo']");
                //if (cert_NameNode == null || CommonUtil.HtmlDecode(cert_NameNode.InnerText.Trim()).Trim() != fullName)
                if (cert_NameNode != null && CommonUtil.HtmlDecode(cert_NameNode.InnerText.Trim()).Trim().ToLower().Replace(" ", " ").Contains(city.ToLower() + " " + state.ToLower() + " "))
                {
                    found = true;
                    this.RunPage.SaveFile(pilotInfoPageHtml, localFilePath, Encoding.GetEncoding("utf-8"));
                    break;
                }
            }
            if (!found)
            {
                throw new CannotFoundException("未找到匹配人. name = " + firstName + " " + lastName);
                //throw new Exception("未找到匹配人. name = " + firstName + " " + lastName);
            }
        }
        public override bool AfterGrabOneCatchException(string pageUrl, Dictionary<string, string> listRow, Exception ex)
        {
            return ex is CannotFoundException;
        }

        #endregion 

        #region 获取网页信息超时时间
        /// <summary>
        /// 获取网页信息超时时间
        /// </summary>
        private int WebRequestTimeout = 40 * 1000;
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
                    throw new Exception("打开页面超时. PageUrl = " + url);
                }
                //等待
                waitCount++;
                Thread.Sleep(SysConfig.WebPageRequestInterval);
            }

            //再增加个等待，等待异步加载的数据
            Thread.Sleep(1000);
            return webBrowser;
        }
        #endregion

        #region InputNameScript
        private void InvokeInputNameScript(WebBrowser webBrowser, string firstName, string lastName, string country, string state)
        {
            webBrowser.Invoke(new InputNameScriptInvokeDelegate(InputNameScript), new object[] { webBrowser, firstName, lastName, country, state });
        }
        private delegate void InputNameScriptInvokeDelegate(WebBrowser webBrowser, string firstName, string lastName, string country, string state);
        private void InputNameScript(WebBrowser webBrowser, string firstName, string lastName, string country, string state)
        {
            AddCheckPageScript(webBrowser);

            webBrowser.Document.GetElementById("ctl00_content_ctl01_txtbxLastName").SetAttribute("value", lastName);
            Thread.Sleep(200);
            webBrowser.Document.GetElementById("ctl00_content_ctl01_txtbxFirstName").SetAttribute("value", firstName);
            Thread.Sleep(200);
            webBrowser.Document.GetElementById("ctl00_content_ctl01_ddlSearchCountry").SetAttribute("value", country);
            Thread.Sleep(200);
            if (state.Length != 0)
            {
                webBrowser.Document.GetElementById("ctl00_content_ctl01_ddlSearchState").SetAttribute("value", state);
                Thread.Sleep(200);
            }
            webBrowser.Document.GetElementById("ctl00_content_ctl01_btnSearch").InvokeMember("click");
            Thread.Sleep(1000);
        }

        private void InvokeClickPilotNameScript(WebBrowser webBrowser, string linkNodeId)
        {
            webBrowser.Invoke(new ClickPilotNameScriptInvokeDelegate(ClickPilotNameScript), new object[] { webBrowser, linkNodeId });
        }
        private delegate void ClickPilotNameScriptInvokeDelegate(WebBrowser webBrowser, string linkNodeId);
        private void ClickPilotNameScript(WebBrowser webBrowser, string linkNodeId)
        {
            AddCheckPageScript(webBrowser);
            webBrowser.Document.GetElementById(linkNodeId).InvokeMember("click");
        }


        private void InvokeAddCheckPageScript(WebBrowser webBrowser)
        {
            webBrowser.Invoke(new AddCheckPageScriptInvokeDelegate(AddCheckPageScript), new object[] { webBrowser});
        }
        private delegate void AddCheckPageScriptInvokeDelegate(WebBrowser webBrowser);
        private void AddCheckPageScript(WebBrowser webBrowser)
        {
            HtmlElement sElement = webBrowser.Document.CreateElement("script");
            IHTMLScriptElement scriptElement = (IHTMLScriptElement)sElement.DomElement;
            scriptElement.text = "function isOldPage(){return 'yes';};";
            webBrowser.Document.Body.AppendChild(sElement);
            string isOldPage = (string)this.RunPage.InvokeDoScriptMethod(webBrowser, "isOldPage", null);
        }
        #endregion

        private string GetIsOldPage(WebBrowser webBrowser)
        {
            try
            {
                string isOldPage = (string)this.RunPage.InvokeDoScriptMethod(webBrowser, "isOldPage", null);
                return isOldPage;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        private void WaitGetPilotInfoPage(WebBrowser webBrowser)
        {
            string isOldPage = this.GetIsOldPage(webBrowser);
            int waitCount = 0;
            while (isOldPage == "yes")
            {
                if (SysConfig.WebPageRequestInterval * waitCount > WebRequestTimeout)
                {
                    throw new Exception("打开页面超时");
                }
                //等待
                waitCount++;
                Thread.Sleep(SysConfig.WebPageRequestInterval);
                isOldPage = this.GetIsOldPage(webBrowser);
            }
            Thread.Sleep(1000);
            InvokeAddCheckPageScript(webBrowser);
        }
    }
}