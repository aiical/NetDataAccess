using HtmlAgilityPack;
using NetDataAccess.Base.Browser;
using NetDataAccess.Base.Common;
using NetDataAccess.Base.Config;
using NetDataAccess.Base.DB;
using NetDataAccess.Base.EnumTypes;
using NetDataAccess.Base.UI;
using NetDataAccess.Base.Web;
using NetDataAccess.Base.Writer;
using NetDataAccess.Extended.Linkedin.Common;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace NetDataAccess.Extended.Linkedin.KeyWordsSearchAtLinkedin
{
    [System.Runtime.InteropServices.ComVisibleAttribute(true)]
    public class GetKeywordsLinkedinDetailPage : ExternalRunWebPage
    {
        public override void WebBrowserHtml_AfterDoNavigate(string pageUrl, Dictionary<string, string> listRow, string tabName)
        {
            this.RunPage.ShowTabPage(tabName);
        }

        public override void WebBrowserHtml_AfterPageLoaded(string pageUrl, Dictionary<string, string> listRow, IWebBrowser webBrowser)
        {
            Thread.Sleep(2000);
            string checkValue = "pv-top-card-v2-section__info mr5";
            string webBrowserUrl = this.RunPage.InvokeGetWebBrowserPageUrl(webBrowser);

            string webText = this.RunPage.InvokeGetPageHtml(webBrowser).ToLower();
            if (!webText.Contains(checkValue))
            {
                throw new GrabRequestException("页面加载地址错误, webBrowserUrl=" + webBrowserUrl + ", pageUrl=" + pageUrl);
            }
            ProcessWebBrowser.AutoScroll(this.RunPage, webBrowser, 3000, 500, 1000, 2000);

            this.ClickAllMoreLinks(webBrowser);

            Random r = new Random(DateTime.Now.Millisecond);
            Thread.Sleep(r.Next(10) * 1000);
        }

        private void ClickAllMoreLinks(IWebBrowser webBrowser)
        {
            AddClickMoreMethod(webBrowser);
        } 
         
        public override void BeforeGrabOne(string pageUrl, Dictionary<string, string> listRow, bool existLocalFile)
        {
            base.BeforeGrabOne(pageUrl, listRow, existLocalFile);
            if (this.CheckNeedLogin())
            {
                if (this.CheckCanLogin())
                {
                    this.Login();
                }
                else
                {
                    Thread.Sleep(1000 * 60 * 10);
                    throw new Exception("没有可用的账号");
                }
            }
        }

        private int _CurrentUserIndex = -1;
        private int _CurrentUserRequestCount = 0;

        private bool CheckNeedLogin()
        {
            if (_CurrentUserIndex < 0)
            {
                this.Logout();
                _CurrentUserIndex++;
                _CurrentUserRequestCount = 0;
                return true;
            }
            else if (_CurrentUserIndex >= _UserInfoList.Count)
            {
                return true;
            }
            else
            {
                string[] userInfo = _UserInfoList[_CurrentUserIndex];
                int requestCountLimit = int.Parse(userInfo[2]);
                if (_CurrentUserRequestCount > requestCountLimit)
                {
                    this.Logout();
                    _CurrentUserIndex++;
                    _CurrentUserRequestCount = 0;
                    return true;
                }
                else
                {
                    _CurrentUserRequestCount++;
                    return false;
                }
            }
        }

        private bool CheckCanLogin()
        {
            if (_CurrentUserIndex < _UserInfoList.Count)
            {
                return true;
            }
            else if(DateTime.Now.Hour == 13)
            {
                //现在是下午1点半到2点之间，那么重新启动爬取
                _CurrentUserIndex = 0;
                return true;
            }
            else
            {
                return false;
            }
        }

        private void Login()
        {
            string pageUrl = "https://www.linkedin.com";
            string tabName = "login";
            IWebBrowser webBrowser = this.RunPage.InvokeShowWebPage(pageUrl, tabName, WebBrowserType.Chromium, false);
            this.RunPage.ShowTabPage(tabName);
            string htmlContent = null;
            int waitCount = 0;
            int timeout = 30000;
            while (htmlContent == null)
            {
                if (SysConfig.WebPageRequestInterval * waitCount > timeout)
                {
                    //超时
                    throw new GrabRequestException("请求Logout页超时. PageUrl = " + pageUrl);
                }
                //等待
                waitCount++;
                Thread.Sleep(SysConfig.WebPageRequestInterval);
                htmlContent = webBrowser.GetDocumentText();
            } 

            string[] userInfo = _UserInfoList[_CurrentUserIndex];

            string userName = userInfo[0].Trim();
            string password = userInfo[1].Trim();

            string inputUserInfoCode = "document.getElementById('login-email').click();document.getElementById('login-email').value = '" + userName + "';document.getElementById('login-password').click();document.getElementById('login-password').value = '" + password + "';document.getElementById('login-submit').disabled=false;";
            webBrowser.AddScriptMethod(inputUserInfoCode);
            Thread.Sleep(3000);

            string submitUserInfoCode = "document.getElementById('login-submit').click();";
            webBrowser.AddScriptMethod(submitUserInfoCode);
            Thread.Sleep(3000);
        }

        private void Logout()
        {
            string pageUrl = "http://www.linkedin.com/m/logout";
            string tabName = "logout";
            IWebBrowser webBrowser = this.RunPage.InvokeShowWebPage(pageUrl, tabName, WebBrowserType.Chromium, false);
            this.RunPage.ShowTabPage(tabName);
            string htmlContent = null;
            int waitCount = 0;
            int timeout = 30000;
            while (htmlContent == null)
            {
                if (SysConfig.WebPageRequestInterval * waitCount > timeout)
                {
                    //超时
                    throw new GrabRequestException("请求Logout页超时. PageUrl = " + pageUrl);
                }
                //等待
                waitCount++;
                Thread.Sleep(SysConfig.WebPageRequestInterval);
                htmlContent = webBrowser.GetDocumentText();
            }
        }

        private List<string[]> _UserInfoList = null;

        public override bool BeforeAllGrab()
        {
            _UserInfoList = new List<string[]>();
            string[] userInfoStrs = this.Parameters.Split(new string[] { ";;;" }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < userInfoStrs.Length; i++)
            {
                string userInfoStr = userInfoStrs[i];

                string[] userInfo = userInfoStr.Split(new string[] { "|||" }, StringSplitOptions.RemoveEmptyEntries);
                _UserInfoList.Add(userInfo);
            }
            return true;
        }



        private void AddClickMoreMethod(IWebBrowser webBrowser)
        {
            string scriptMethodCodeA = "$('button[class=\"pv-profile-section__card-action-bar pv-skills-section__additional-skills artdeco-container-card-action-bar\"]').click();" ;
            string scriptMethodCodeB = "$('button[class=\"pv-profile-section__see-more-inline pv-profile-section__text-truncate-toggle link\"]').click();";

            this.RunPage.InvokeAddScriptMethod(webBrowser, scriptMethodCodeA);
            Thread.Sleep(3000);
            this.RunPage.InvokeAddScriptMethod(webBrowser, scriptMethodCodeB); 
            Thread.Sleep(3000);
        }


        public override bool AfterAllGrab(IListSheet listSheet)
        {
            try
            {
                this.GetUserInfoInPages(listSheet);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return true;
        }

        private void GetUserInfoInPages(IListSheet listSheet)
        {
            ExcelWriter ew = this.GetExcelBaseWriter();
            ExcelWriter gzjlEw = this.GetExcelExpWriter();
            ExcelWriter jyjlEw = this.GetExcelEduWriter();
            ExcelWriter zgrzEw = this.GetExcelCertificationWriter();
            ExcelWriter yynlEw = this.GetExcelLanguageWriter();

            int rowCount = listSheet.GetListDBRowCount();
            for (int i = 0; i < rowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string keywords = row["keywords"];
                    string pageUrl = row["detailPageUrl"];

                    HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);

                    HtmlNode xmNode = htmlDoc.DocumentNode.SelectSingleNode("//h1[contains(@class,\"pv-top-card-section__name \")]");
                    string xm = CommonUtil.HtmlDecode(xmNode.InnerText).Trim();

                    HtmlNode mqgzNode = htmlDoc.DocumentNode.SelectSingleNode("//h2[contains(@class,\"pv-top-card-section__headline \")]");
                    string mqgz = mqgzNode == null ? "" : CommonUtil.HtmlDecode(mqgzNode.InnerText).Trim();

                    HtmlNode dqNode = htmlDoc.DocumentNode.SelectSingleNode("//h3[contains(@class,\"pv-top-card-section__location \")]");
                    string dq = dqNode == null ? "" : CommonUtil.HtmlDecode(dqNode.InnerText).Trim();

                    HtmlNode gsNode = htmlDoc.DocumentNode.SelectSingleNode("//span[contains(@class,\"pv-top-card-v2-section__company-name \")]");
                    string gs = gsNode == null ? "" : CommonUtil.HtmlDecode(gsNode.FirstChild.InnerText).Trim();

                    HtmlNode xxNode = htmlDoc.DocumentNode.SelectSingleNode("//span[contains(@class,\"pv-top-card-v2-section__school-name \")]");
                    string xx = xxNode == null ? "" : CommonUtil.HtmlDecode(xxNode.FirstChild.InnerText).Trim();

                    JObject infoJson = null;
                    HtmlNodeCollection codeNodes = htmlDoc.DocumentNode.SelectNodes("//code");
                    if (codeNodes != null)
                    {
                        foreach (HtmlNode codeNode in codeNodes)
                        {
                            string text = CommonUtil.HtmlDecode(codeNode.InnerText);
                            if (text.Contains("positionGroupView"))
                            {
                                infoJson = JObject.Parse(text);
                                break;
                            }
                        }
                    }

                    StringBuilder gzjl = new StringBuilder();
                    HtmlNodeCollection gzjlLiNodes = htmlDoc.DocumentNode.SelectNodes("//section[@id=\"experience-section\"]/ul/div/li");
                    if (gzjlLiNodes != null)
                    {
                        foreach (HtmlNode gzjlLiNode in gzjlLiNodes)
                        {
                            string gzjlId = gzjlLiNode.GetAttributeValue("id", "");

                            HtmlNode gzjlNode = gzjlLiNode.SelectSingleNode("./a/div[contains(@class, \"pv-entity__summary-info \")]");
                            if (gzjlNode != null)
                            {

                                HtmlNode zwNode = gzjlNode.SelectSingleNode("./h3");
                                string zw = zwNode == null ? "" : CommonUtil.HtmlDecode(zwNode.InnerText.Trim());

                                HtmlNode rzgsNode = gzjlNode.SelectSingleNode("./h4/span[@class=\"pv-entity__secondary-title\"]");
                                string rzgs = rzgsNode == null ? "" : CommonUtil.HtmlDecode(rzgsNode.InnerText.Trim());

                                HtmlNode rzrqNode = gzjlNode.SelectSingleNode("./div[@class=\"display-flex\"]/h4[contains(@class, \"pv-entity__date-range\")]/span[last()]");
                                string rzrq = rzrqNode == null ? "" : CommonUtil.HtmlDecode(rzrqNode.InnerText.Trim());

                                HtmlNode rzscNode = gzjlNode.SelectSingleNode("./div[@class=\"display-flex\"]/h4[last()]/span[@class=\"pv-entity__bullet-item-v2\"]");
                                string rzsc = rzscNode == null ? "" : CommonUtil.HtmlDecode(rzscNode.InnerText.Trim());

                                HtmlNode szdqNode = gzjlNode.SelectSingleNode("./h4[contains(@class, \"pv-entity__location \")]/span[last()]");
                                string szdq = szdqNode == null ? "" : CommonUtil.HtmlDecode(szdqNode.InnerText.Trim());

                                string gzjlms = GetExpDescription(infoJson, gzjlId).Replace("\r\n", " ");

                                gzjl.Append(zw + ", " + rzgs + ", " + rzrq + ", " + rzsc + ", " + szdq + ", " + gzjlms + ". ");

                                /*
                                HtmlNodeCollection extraNodes = gzjlNode.SelectNodes("./div[contains(@class, \"pv-entity__extra-details\")]/p");
                                StringBuilder gzjlmsBuilder = new StringBuilder();
                                if (extraNodes != null)
                                {
                                    foreach (HtmlNode extraNode in extraNodes)
                                    {
                                        gzjlmsBuilder.Append(extraNode != null ? "" : CommonUtil.HtmlDecode(extraNode.InnerText.Trim()).Replace("\r\n", " "));
                                    }
                                    gzjl.Append(gzjlmsBuilder.ToString());
                                }
                                */

                                gzjl.AppendLine();
                                this.AddExpRow(gzjlEw, xm, pageUrl, zw, rzgs, rzrq, rzsc, szdq, gzjlms);
                            }
                            else
                            {
                                HtmlNode gzjlMultiNode = gzjlLiNode.SelectSingleNode("./a/div[contains(@class, \"pv-entity__company-details\")]");

                                HtmlNode rzgsNode = gzjlMultiNode.SelectSingleNode("./div[contains(@class, \"pv-entity__company-summary-info\")]/h3/span[last()]");
                                string rzgs = rzgsNode == null ? "" : CommonUtil.HtmlDecode(rzgsNode.InnerText.Trim());

                                //HtmlNode rzscNode = gzjlMultiNode.SelectSingleNode("./div[contains(@class, \"pv-entity__company-summary-info\")]/h4/span[last()]");
                                //string rzsc = rzscNode == null ? "" : CommonUtil.HtmlDecode(rzscNode.InnerText.Trim());

                                HtmlNodeCollection roleNodes = gzjlLiNode.SelectNodes("./ul/li/div/div/div[contains(@class, \"pv-entity__role-details-container\")]");
                                if (roleNodes != null)
                                {
                                    foreach (HtmlNode roleNode in roleNodes)
                                    {
                                        HtmlNode zwNode = roleNode.SelectSingleNode("./div[contains(@class, \"pv-entity__summary-info-v2\")]/h3/span[last()]");
                                        string zw = zwNode == null ? "" : CommonUtil.HtmlDecode(zwNode.InnerText.Trim());

                                        HtmlNode rzrqNode = roleNode.SelectSingleNode("./div[contains(@class, \"pv-entity__summary-info-v2\")]/div[@class=\"display-flex\"]/h4[contains(@class, \"pv-entity__date-range\")]/span[last()]");
                                        string rzrq = rzrqNode == null ? "" : CommonUtil.HtmlDecode(rzrqNode.InnerText.Trim());

                                        HtmlNode rzscNode = roleNode.SelectSingleNode("./div[contains(@class, \"pv-entity__summary-info-v2\")]/div[@class=\"display-flex\"]/h4[last()]/span[@class=\"pv-entity__bullet-item-v2\"]");
                                        string rzsc = rzscNode == null ? "" : CommonUtil.HtmlDecode(rzscNode.InnerText.Trim());

                                        HtmlNode szdqNode = roleNode.SelectSingleNode("./h4[contains(@class, \"pv-entity__location \")]/span[last()]");
                                        string szdq = szdqNode == null ? "" : CommonUtil.HtmlDecode(szdqNode.InnerText.Trim());

                                        HtmlNodeCollection extraNodes = roleNode.SelectNodes("./div[contains(@class, \"pv-entity__extra-details\")]/p");
                                        StringBuilder gzjlmsBuilder = new StringBuilder();
                                        if (extraNodes != null)
                                        {
                                            foreach (HtmlNode extraNode in extraNodes)
                                            {
                                                gzjlmsBuilder.Append(CommonUtil.HtmlDecode(extraNode.InnerText.Trim()).Replace("\r\n", " "));
                                            }
                                        }

                                        gzjl.Append(zw + ", " + rzgs + ", " + rzrq + ", " + rzsc + ", " + szdq + ", " + gzjlmsBuilder.ToString() + ". ");
                                        this.AddExpRow(gzjlEw, xm, pageUrl, zw, rzgs, rzrq, rzsc, szdq, gzjlmsBuilder.ToString());
                                    }
                                }

                            }
                        }
                    }

                    StringBuilder jyjl = new StringBuilder();
                    HtmlNodeCollection jyjlANodes = htmlDoc.DocumentNode.SelectNodes("//section[@id=\"education-section\"]/ul/li/a");
                    HtmlNodeCollection jyjlNodes = null;
                    if (jyjlANodes == null)
                    {
                        jyjlNodes = htmlDoc.DocumentNode.SelectNodes("//section[@id=\"education-section\"]/ul/li/div");
                    }
                    else
                    {
                        jyjlNodes = htmlDoc.DocumentNode.SelectNodes("//section[@id=\"education-section\"]/ul/li");
                    }
                    if (jyjlNodes != null)
                    {
                        foreach (HtmlNode jyjlNode in jyjlNodes)
                        {
                            HtmlNode summaryNode = jyjlNode.SelectSingleNode("./a/div[contains(@class, \"pv-entity__summary-info \")]");
                            HtmlNode xuexiaoNode;
                            try
                            {
                                xuexiaoNode = summaryNode.SelectSingleNode("./div/h3");
                            }
                            catch (Exception ex)
                            {
                                throw ex;
                            }
                            string xuexiao = xuexiaoNode == null ? "" : CommonUtil.HtmlDecode(xuexiaoNode.InnerText).Trim();

                            HtmlNode xueweiNode = summaryNode.SelectSingleNode("./div/p[contains(@class, \"pv-entity__degree-name\")]/span[contains(@class, \"pv-entity__comma-item\")]");
                            string xuewei = xueweiNode == null ? "" : CommonUtil.HtmlDecode(xueweiNode.InnerText).Trim();

                            HtmlNode zhuanyeNode = summaryNode.SelectSingleNode("./div/p[contains(@class, \"pv-entity__fos\")]/span[contains(@class, \"pv-entity__comma-item\")]");
                            string zhuanye = zhuanyeNode == null ? "" : CommonUtil.HtmlDecode(zhuanyeNode.InnerText).Trim();

                            HtmlNode chengjiNode = summaryNode.SelectSingleNode("./div/p[contains(@class, \"pv-entity__grade\")]/span[contains(@class, \"pv-entity__comma-item\")]");
                            string chengji = chengjiNode == null ? "" : CommonUtil.HtmlDecode(chengjiNode.InnerText).Trim();

                            HtmlNode shijianNode = summaryNode.SelectSingleNode("./p[contains(@class, \"pv-entity__dates\")]/span[last()]");
                            string shijian = shijianNode == null ? "" : CommonUtil.HtmlDecode(shijianNode.InnerText).Trim();

                            jyjl.Append(xuexiao + ", " + xuewei + ", " + zhuanye + ", " + chengji + ", " + shijian + ". ");

                            HtmlNodeCollection extraNodes = jyjlNode.SelectNodes("./div[contains(@class, \"pv-entity__extra-details\")]/p");
                            StringBuilder miaoshuBuilder = new StringBuilder();
                            if (extraNodes != null)
                            {
                                foreach (HtmlNode extraNode in extraNodes)
                                {
                                    string buchong = CommonUtil.HtmlDecode(extraNode.InnerText.Trim()).Replace("\r\n", " ");
                                    miaoshuBuilder.Append(buchong);
                                }
                                jyjl.Append(miaoshuBuilder.ToString());
                            }

                            this.AddEduRow(jyjlEw, xm, pageUrl, xx, xuewei, zhuanye, chengji, shijian, miaoshuBuilder.ToString());
                            jyjl.AppendLine();
                        }
                    }

                    StringBuilder jn = new StringBuilder();
                    HtmlNodeCollection jnNodes = htmlDoc.DocumentNode.SelectNodes("//section[contains(@class,\"pv-skill-categories-section\")]/ol/li");
                    if (jnNodes != null)
                    {
                        foreach (HtmlNode jnNode in jnNodes)
                        {
                            HtmlNode jnmcNode = jnNode.SelectSingleNode("./div/p/a/span");
                            if (jnmcNode != null)
                            {
                                string jnmc = CommonUtil.HtmlDecode(jnmcNode.InnerText).Trim();

                                HtmlNode jnrenkeNode = jnNode.SelectSingleNode("./div/a/span[contains(@class, \"pv-skill-category-entity__endorsement-count\")]");
                                string jnrenke = jnrenkeNode == null ? "" : CommonUtil.HtmlDecode(jnrenkeNode.InnerText).Trim();

                                jn.Append(jnmc + (jnrenke.Length == 0 ? "" : (", " + jnrenke)) + "; ");
                            }
                            else
                            {
                                jnmcNode = jnNode.SelectSingleNode("./div/p");

                                string jnmc = CommonUtil.HtmlDecode(jnmcNode.InnerText).Trim();
                                jn.Append(jnmc + "; ");
                            }
                        }
                    }
                    HtmlNodeCollection otherJnNodes = htmlDoc.DocumentNode.SelectNodes("//section[contains(@class,\"pv-skill-categories-section\")]/div[@id=\"skill-categories-expanded\"]/div/ol/li");
                    if (otherJnNodes != null)
                    {
                        foreach (HtmlNode jnNode in otherJnNodes)
                        {
                            HtmlNode jnmcNode = jnNode.SelectSingleNode("./div/p/a/span");
                            if (jnmcNode != null)
                            {
                                string jnmc = CommonUtil.HtmlDecode(jnmcNode.InnerText).Trim();

                                HtmlNode jnrenkeNode = jnNode.SelectSingleNode("./div/a/span[contains(@class, \"pv-skill-category-entity__endorsement-count\")]");
                                string jnrenke = jnrenkeNode == null ? "" : CommonUtil.HtmlDecode(jnrenkeNode.InnerText).Trim();

                                jn.Append(jnmc + (jnrenke.Length == 0 ? "" : (", " + jnrenke)) + "; ");
                            }
                            else
                            {
                                jnmcNode = jnNode.SelectSingleNode("./div/p");

                                string jnmc = CommonUtil.HtmlDecode(jnmcNode.InnerText).Trim();
                                jn.Append(jnmc + "; ");
                            }
                        }
                    }

                    string szxm = "";
                    string zgrz = "";
                    string cbzp = "";
                    string yynl = "";
                    string ryjx = "";
                    string cyzz = "";
                    string sxkc = "";
                    string cscj = "";
                    string zlfm = "";
                    HtmlNodeCollection grcjNodes = htmlDoc.DocumentNode.SelectNodes("//section[contains(@class, \"pv-accomplishments-section\")]/div/section/div");
                    if (grcjNodes != null)
                    {
                        foreach (HtmlNode grcjNode in grcjNodes)
                        {
                            HtmlNode grcjmcNode = grcjNode.SelectSingleNode("./h3");
                            string grcjmc = CommonUtil.HtmlDecode(grcjmcNode.InnerText).Trim();

                            StringBuilder grcjnr = new StringBuilder();
                            HtmlNodeCollection grcjnrNodes = grcjNode.SelectNodes("./div/ul/li");
                            if (grcjnrNodes != null)
                            {
                                if (grcjnrNodes.Count == 1)
                                {
                                    string value = CommonUtil.HtmlDecode(grcjnrNodes[0].InnerText).Trim();
                                    grcjnr.Append(value);
                                    switch (grcjmc)
                                    {
                                        case "资格认证":
                                        case "Certifications":
                                        case "Certification":
                                            this.AddCertificationRow(zgrzEw, xm, pageUrl, value);
                                            break;
                                        case "语言能力":
                                        case "Language":
                                        case "Languages":
                                            this.AddLanguageRow(yynlEw, xm, pageUrl, value);
                                            break;
                                    }
                                }
                                else
                                {
                                    foreach (HtmlNode grcjnrNode in grcjnrNodes)
                                    {
                                        string value = CommonUtil.HtmlDecode(grcjnrNode.InnerText).Trim();
                                        grcjnr.Append(value + "; ");
                                        switch (grcjmc)
                                        {
                                            case "资格认证":
                                            case "Certifications":
                                            case "Certification":
                                                this.AddCertificationRow(zgrzEw, xm, pageUrl, value);
                                                break;
                                            case "语言能力":
                                            case "Language":
                                            case "Languages":
                                                this.AddLanguageRow(yynlEw, xm, pageUrl, value);
                                                break;
                                        }
                                    }
                                }
                            }
                            switch (grcjmc)
                            {
                                case "所做项目":
                                case "Project":
                                case "Projects":
                                    szxm = grcjnr.ToString();
                                    break;
                                case "资格认证":
                                case "Certifications":
                                case "Certification":
                                    zgrz = grcjnr.ToString();
                                    break;
                                case "出版作品":
                                case "Publications":
                                case "Publication":
                                    cbzp = grcjnr.ToString();
                                    break;
                                case "语言能力":
                                case "Language":
                                case "Languages":
                                    yynl = grcjnr.ToString();
                                    break;
                                case "荣誉奖项":
                                case "Honors & Awards":
                                case "Honor & Award":
                                    ryjx = grcjnr.ToString();
                                    break;
                                case "参与组织":
                                case "Organization":
                                case "Organizations":
                                    cyzz = grcjnr.ToString();
                                    break;
                                case "所学课程":
                                case "Courses":
                                case "Course":
                                    sxkc = grcjnr.ToString();
                                    break;
                                case "测试成绩":
                                case "Test Score":
                                case "Test Scores":
                                    cscj = grcjnr.ToString();
                                    break;
                                case "专利发明":
                                case "Patent":
                                case "Patents":
                                    zlfm = grcjnr.ToString();
                                    break;
                                default:
                                    MessageBox.Show("个人成就: " + grcjmc);
                                    break;
                            }
                        }
                    }

                    Dictionary<string, string> p2vs = new Dictionary<string, string>();
                    p2vs.Add("搜索关键词", keywords);
                    p2vs.Add("名字", xm);
                    p2vs.Add("目前工作", mqgz);
                    p2vs.Add("地区", dq);
                    p2vs.Add("公司", gs);
                    p2vs.Add("学校", xx);
                    p2vs.Add("url", pageUrl);
                    p2vs.Add("工作经历", gzjl.ToString());
                    p2vs.Add("教育经历", jyjl.ToString());
                    p2vs.Add("技能", jn.ToString());
                    p2vs.Add("所做项目", szxm);
                    p2vs.Add("资格认证", zgrz);
                    p2vs.Add("出版作品", cbzp);
                    p2vs.Add("语言能力", yynl);
                    p2vs.Add("荣誉奖项", ryjx);
                    p2vs.Add("参与组织", cyzz);
                    p2vs.Add("测试成绩", cscj);
                    p2vs.Add("所学课程", sxkc);
                    p2vs.Add("专利发明", zlfm);
                    ew.AddRow(p2vs);
                }
            }
            ew.SaveToDisk();
            gzjlEw.SaveToDisk();
            jyjlEw.SaveToDisk();
            zgrzEw.SaveToDisk();
            yynlEw.SaveToDisk();
        }

        private string GetExpDescription(JObject infoJson, string gzglId)
        {
            JArray elements = infoJson.SelectToken("included") as JArray;
            if (elements != null)
            {
                for (int i = 0; i < elements.Count; i++)
                {
                    JObject elementObj = elements[i] as JObject;
                    JToken companyObj = elementObj.SelectToken("companyName");
                    if (companyObj != null)
                    { 
                        string entityUrn = elementObj.SelectToken("entityUrn").ToString();
                        if (entityUrn.Contains("," + gzglId + ")"))
                        {
                            JToken descriptionObj = elementObj.SelectToken("description");
                            return descriptionObj == null ? "" : descriptionObj.ToString();
                        }
                    }
                }
            }
            return "";
        }

        private ExcelWriter GetExcelBaseWriter()
        {
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                "搜索关键词",
                "名字", 
                "目前工作", 
                "地区",  
                "公司", 
                "学校", 
                "url", 
                "工作经历",  
                "教育经历", 
                "技能", 
                "所做项目", 
                "资格认证", 
                "出版作品", 
                "语言能力", 
                "荣誉奖项", 
                "参与组织", 
                "所学课程", 
                "测试成绩", 
                "专利发明"});
            string resultFilePath = Path.Combine(exportDir, "Linkedin个人基本信息.xlsx");

            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic);
            return resultEW;
        }

        private void AddExpRow(ExcelWriter ew, string xm, string url, string zw, string rzgs, string rgsj, string rzsc, string szdq, string ms)
        {
            Dictionary<string, string> row = new Dictionary<string, string>();
            row.Add("url", url);
            row.Add("名字", xm);
            row.Add("职位", zw);
            row.Add("任职公司", rzgs);
            row.Add("任职时间", rgsj);
            row.Add("任职时长", rzsc);
            row.Add("所在地区", szdq);
            row.Add("描述", ms);
            ew.AddRow(row);
        }

        private ExcelWriter GetExcelExpWriter()
        {
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                "url",
                "名字", 
                "职位", 
                "任职公司", 
                "任职时间", 
                "任职时长", 
                "所在地区", 
                "描述"});
            string resultFilePath = Path.Combine(exportDir, "Linkedin个人工作经验.xlsx");

            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic);
            return resultEW;
        }

        private void AddEduRow(ExcelWriter ew, string xm, string url, string xx, string xw, string zy, string cj, string xxsj, string ms)
        {
            Dictionary<string, string> row = new Dictionary<string, string>();
            row.Add("url", url);
            row.Add("名字", xm);
            row.Add("学校", xx);
            row.Add("学位", xw);
            row.Add("专业", zy);
            row.Add("成绩", cj);
            row.Add("学习时间", xxsj);
            row.Add("描述", ms);
            ew.AddRow(row);
        }

        private ExcelWriter GetExcelEduWriter()
        {
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                "url",
                "名字", 
                "学校", 
                "学位",  
                "专业",  
                "成绩",  
                "学习时间",   
                "描述"});
            string resultFilePath = Path.Combine(exportDir, "Linkedin个人教育经历.xlsx");

            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic);
            return resultEW;
        }

        private void AddCertificationRow(ExcelWriter ew, string xm, string url, string certification)
        {
            Dictionary<string, string> row = new Dictionary<string, string>();
            row.Add("url", url);
            row.Add("名字", xm);
            row.Add("资格认证", certification);
            ew.AddRow(row);
        }

        private ExcelWriter GetExcelCertificationWriter()
        {
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                "url",
                "名字", 
                "资格认证"});
            string resultFilePath = Path.Combine(exportDir, "Linkedin资格认证.xlsx");

            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic);
            return resultEW;
        }

        private void AddLanguageRow(ExcelWriter ew,string xm,string url, string language)
        {
            Dictionary<string, string> row = new Dictionary<string, string>();
            row.Add("url", url);
            row.Add("名字", xm);
            row.Add("语言", language);
            ew.AddRow(row);
        }

        private ExcelWriter GetExcelLanguageWriter()
        {
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                "url",
                "名字", 
                "语言"});
            string resultFilePath = Path.Combine(exportDir, "Linkedin语言能力.xlsx");

            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic);
            return resultEW;
        }

    }
}
