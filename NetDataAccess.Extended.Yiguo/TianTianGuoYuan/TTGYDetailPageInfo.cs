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

namespace NetDataAccess.Extended.Yiguo
{
    /// <summary>
    /// 天天果园
    /// 获取商品信息
    /// </summary>
    public class TTGYDetailPageInfo : CustomProgramBase
    {
        #region 入口函数
        public bool Run(string parameters, IListSheet listSheet)
        {
            GetPriceJson(listSheet, parameters);
            return this.GenerateDetailPageInfo(listSheet);
        }
        #endregion

        #region 从网站逐个获取价格信息
        private void GetPriceJson(IListSheet listSheet, string parameters)
        {
            //允许跳转到查询页面的次数，有时会出现跳转至登录页面的情况
            const int allowGoToQueryPageCount = 10;

            int goToQueryPageErrorCount = 0;

            string pageUrl = parameters; 

            string currentUrl = "";

            while (currentUrl != pageUrl && goToQueryPageErrorCount < allowGoToQueryPageCount)
            {
                //加载网页
                this.ShowWebPage(pageUrl); 

                currentUrl = WebBrowserMain.Url.ToString();
            }

            if (currentUrl != pageUrl)
            {
                throw new Exception("无法定位到查询页面.");
            }
            else
            {
                string sourceFileDir = this.RunPage.GetDetailSourceFileDir();
                for (int i = 0; i < listSheet.RowCount; i++)
                {
                    Dictionary<string, string> row = listSheet.GetRow(i);
                    string productCode = row["productCode"];
                    string localCommmentFilePath = this.RunPage.GetFilePath(productCode, sourceFileDir);
                    if (!File.Exists(localCommmentFilePath))
                    {
                        //通过ajax调用
                        this.InvokeGetMyData(WebBrowserMain, productCode);

                        //轮询次数
                        int waitCount = 0;

                        //记录查询到的返回值
                        string resultValue = "";

                        //存在异步加载数据的情况，此处用轮询获取查询到的数据
                        while (resultValue == null || !resultValue.StartsWith(productCode + ","))
                        {
                            if (SysConfig.WebPageRequestInterval * waitCount > WebRequestTimeout)
                            {
                                //超时
                                string errorInfo = "获取商品评价超时,没有获取到返回值! productSysNo=" + productCode;
                                this.RunPage.InvokeAppendLogText(errorInfo, Base.EnumTypes.LogLevelType.System, true);
                                throw new Exception(errorInfo);
                            }
                            waitCount++;
                            Thread.Sleep(SysConfig.WebPageRequestInterval);
                            try
                            {
                                resultValue = this.InvokeReadMyData(WebBrowserMain);
                            }
                            catch (Exception dex)
                            {
                                //超时
                                string errorInfo = "调用异常, InvokeReadMyData! productSysNo=" + productCode;
                                this.RunPage.InvokeAppendLogText(errorInfo, Base.EnumTypes.LogLevelType.System, true);
                                throw new Exception(errorInfo, dex);
                            }
                        }

                        this.RunPage.InvokeAppendLogText(resultValue, Base.EnumTypes.LogLevelType.System, true);

                        CommonUtil.CreateFileDirectory(localCommmentFilePath);
                        StreamWriter sw = null;
                        try
                        {
                            sw = new StreamWriter(localCommmentFilePath, false, Encoding.UTF8);
                            sw.Write(resultValue);
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
            }
        }
        #endregion

        #region 浏览器控件
        /// <summary>
        /// 浏览器控件
        /// </summary>
        private WebBrowser WebBrowserMain = null;
        #endregion 

        #region 获取网页信息超时时间
        /// <summary>
        /// 获取网页信息超时时间
        /// </summary>
        private int WebRequestTimeout = 20 * 1000;
        #endregion
        
        #region 显示网页
        private void ShowWebPage(string url)
        {
            this.RunPage.InvokeShowWebPage(url, "");
            int waitCount = 0;
            while (!this.RunPage.CheckIsComplete(""))
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

            WebBrowserMain = this.RunPage.GetWebBrowserByName("");

            this.InvokeAddMyScript(WebBrowserMain);

            //再增加个等待，等待异步加载的数据
            Thread.Sleep(1000);
        }
        #endregion

        #region 注入js代码到网页
        private void InvokeAddMyScript(WebBrowser webBrowser)
        {
            webBrowser.Invoke(new AddMyScriptInvokeDelegate(AddMyScript), new object[] { webBrowser, "" });
        }
        private delegate void AddMyScriptInvokeDelegate(WebBrowser webBrowser, string p1);
        private void AddMyScript(WebBrowser webBrowser, string p1)
        {
            {
                HtmlElement sElement = webBrowser.Document.CreateElement("script");
                IHTMLScriptElement scriptElement = (IHTMLScriptElement)sElement.DomElement;
                scriptElement.text = "var tempProductCode = \"\";"
                                    + "function getMyData(id){" 
                                    + "$('#product_id').val(id); "
                                    + "getCommentRate();" 
                                    +"}";
                webBrowser.Document.Body.AppendChild(sElement);
            }
            {
                HtmlElement sElement = webBrowser.Document.CreateElement("script");
                IHTMLScriptElement scriptElement = (IHTMLScriptElement)sElement.DomElement;
                scriptElement.text = "function readMyData(id){"
                                    + "var str = (tempProductCode + \",\");"
                                    + "str += ($('#comment_total').text() + \",\");"
                                    + "str += ($('#comment_total_good').text() + \",\");"
                                    + "str += ($('#comment_total_normal').text() + \",\");"
                                    + "str += $('#comment_total_bad').text();"  
                                    + "return str;"
                                    + "}";
                webBrowser.Document.Body.AppendChild(sElement);
            }
            {
                HtmlElement sElement = webBrowser.Document.CreateElement("script");
                IHTMLScriptElement scriptElement = (IHTMLScriptElement)sElement.DomElement;
                scriptElement.text = @"function getCommentRate() {
                                        $.ajax({
                                            type: 'POST',
                                            url: '/ajax/comment/pRate',
                                            dataType: 'json',
                                            data: {
                                                id: $('#product_id').val()
                                            },
                                            success: function(data) {
                                                if (data.code == 200) {
                                                    tempProductCode = $('#product_id').val();
                                                    $('#comment_total_top').text('(' + data.msg.num.total + ')');
                                                    $('#grade_good_another_left').text(data.msg.good + '%');
                                                    $('#grade_good_another').animate({width: data.msg.good + '%'}, 1000);
                                                    $('#good_rate_another').text(data.msg.good + '%');
                                                    $('#grade_normal_another').animate({width: data.msg.normal + '%'}, 1000);
                                                    $('#normal_rate_another').text(data.msg.normal + '%');
                                                    $('#grade_bad_another').animate({width: data.msg.bad + '%'}, 1000);
                                                    $('#bad_rate_another').text(data.msg.bad + '%');
                                                    $('#comment_total').text(data.msg.num.total);
                                                    $('#comment_total_good').text(data.msg.num.good);
                                                    $('#comment_total_normal').text(data.msg.num.normal);
                                                    $('#comment_total_bad').text(data.msg.num.bad);

                                                    getComment('', 0, data.msg.num.total);
                                                }
                                            }
                                        });
                                    }";
                webBrowser.Document.Body.AppendChild(sElement);
                /*
                 function getCommentRate() {
                $.ajax({
                    type: 'POST',
                    url: '/ajax/comment/pRate',
                    dataType: 'json',
                    data: {
                        id: $("#product_id").val()
                    },
                    success: function(data) {
                        if (data.code == 200) {
                            $("#comment_total_top").text('(' + data.msg.num.total + ')');
                            $("#grade_good_another_left").text(data.msg.good + '%');
                            $("#grade_good_another").animate({width: data.msg.good + '%'}, 1000);
                            $("#good_rate_another").text(data.msg.good + '%');
                            $("#grade_normal_another").animate({width: data.msg.normal + '%'}, 1000);
                            $("#normal_rate_another").text(data.msg.normal + '%');
                            $("#grade_bad_another").animate({width: data.msg.bad + '%'}, 1000);
                            $("#bad_rate_another").text(data.msg.bad + '%');
                            $("#comment_total").text(data.msg.num.total);
                            $("#comment_total_good").text(data.msg.num.good);
                            $("#comment_total_normal").text(data.msg.num.normal);
                            $("#comment_total_bad").text(data.msg.num.bad);

                            getComment('', 0, data.msg.num.total);
                        }
                    }
                });
            }
                 */
            }
        }
        #endregion

        #region 读取获取到的数据
        private void InvokeGetMyData(WebBrowser webBrowser, string code)
        {
            webBrowser.Invoke(new GetMyDataInvokeDelegate(GetMyData), new object[] { webBrowser, code });
        }
        private delegate void GetMyDataInvokeDelegate(WebBrowser webBrowser, string code);
        private void GetMyData(WebBrowser webBrowser, string code)
        {
            webBrowser.Document.InvokeScript("getMyData", new string[] { code });
        }
        #endregion

        #region 读取获取到的数据
        private string InvokeReadMyData(WebBrowser webBrowser)
        {
            string districts = (string)webBrowser.Invoke(new ReadMyDataInvokeDelegate(ReadMyData), new object[] { webBrowser });
            return districts;
        }
        private delegate string ReadMyDataInvokeDelegate(WebBrowser webBrowser);
        private string ReadMyData(WebBrowser webBrowser)
        {
            return (string)webBrowser.Document.InvokeScript("readMyData");
        }
        #endregion

        #region 生成商品详情信息文件
        private bool GenerateDetailPageInfo(IListSheet listSheet)
        {
            bool succeed = true;
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("商品名称", 0);
            resultColumnDic.Add("价格", 1);
            resultColumnDic.Add("分类", 2);
            resultColumnDic.Add("分类编码", 3);
            resultColumnDic.Add("规格", 4);
            resultColumnDic.Add("评论数", 5);
            resultColumnDic.Add("好评", 6);
            resultColumnDic.Add("中评", 7);
            resultColumnDic.Add("差评", 8);
            resultColumnDic.Add("好评度", 9);
            resultColumnDic.Add("url", 10);
            resultColumnDic.Add("参考价", 11);
            resultColumnDic.Add("商品编码", 12);
            resultColumnDic.Add("备注", 13);
            resultColumnDic.Add("甜度", 14);
            resultColumnDic.Add("城市", 15);


            Dictionary<string, string> resultColumnFormat = new Dictionary<string, string>();
            resultColumnFormat.Add("价格", "#,##0.00");
            resultColumnFormat.Add("评论数", "#,##0");
            resultColumnFormat.Add("好评", "#,##0");
            resultColumnFormat.Add("中评", "#,##0");
            resultColumnFormat.Add("差评", "#,##0");
            resultColumnFormat.Add("好评度", "0.00%");
            resultColumnFormat.Add("参考价", "#,##0.00");
            resultColumnFormat.Add("甜度", "#,##0");

            string readDetailDir = this.RunPage.GetReadFileDir();
            string exportDir = this.RunPage.GetExportDir();
            string resultFilePath = Path.Combine(exportDir, "天天果园商品详情" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, resultColumnFormat);

            GenerateDetailPageInfo(listSheet, pageSourceDir, resultEW); 

            resultEW.SaveToDisk(); 

            return succeed;
        }
        #endregion

        #region 生成商品详情信息文件
        private void GenerateDetailPageInfo(IListSheet listSheet, string pageSourceDir, ExcelWriter resultEW)
        {
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string pageUrl = row["detailPageUrl"];
                    string productCode = row["productCode"];
                    string city = row["city"];
                    string localFilePath = this.RunPage.GetFilePath(pageUrl, pageSourceDir);
                    string localCommentFilePath = this.RunPage.GetFilePath(productCode, pageSourceDir);
                    string productName = row["productName"];
                    string productCurrentPrice = row["productCurrentPrice"];
                    string productOldPrice = row["productOldPrice"];
                    string categoryCode = row["categoryCode"];
                    string categoryName = row["categoryName"];
                    string standard = row["standard"];
                    int totalCommentCount = 0;
                    int hCommentCount = 0;
                    int mCommentCount = 0;
                    int lCommentCount = 0;
                    Nullable<decimal> hPer = null;
                    string note = "";
                    Nullable<int> sweetness = null;


                    TextReader htmlTr = null;
                    TextReader commentTr = null;

                    try
                    {

                        commentTr = new StreamReader(localCommentFilePath);
                        string commentStr = commentTr.ReadToEnd();
                        string[] commentCountStrs = commentStr.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                        totalCommentCount = int.Parse(commentCountStrs[1].Trim());
                        hCommentCount = int.Parse(commentCountStrs[2].Trim());
                        mCommentCount = int.Parse(commentCountStrs[3].Trim());
                        lCommentCount = int.Parse(commentCountStrs[4].Trim());

                        hPer = totalCommentCount == 0 ? null : (Nullable<decimal>)(((decimal)hCommentCount) / (decimal)totalCommentCount);
                        
                        htmlTr = new StreamReader(localFilePath);
                        string webPageHtml = htmlTr.ReadToEnd();
                        HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                        htmlDoc.LoadHtml(webPageHtml);

                        HtmlNode oldPriceNode = htmlDoc.DocumentNode.SelectSingleNode("//*[@id=\"jq-old-price\"]");
                        if (oldPriceNode != null)
                        {
                            productOldPrice = oldPriceNode.InnerText.Trim().Substring(1);
                        }

                        /*
                        HtmlNode totalCommentNode = htmlDoc.DocumentNode.SelectSingleNode("//*[@id=\"comment_total\"]");
                        HtmlNode goodCommentNode = htmlDoc.DocumentNode.SelectSingleNode("//*[@id=\"comment_total_good\"]");
                        HtmlNode normalCommentNode = htmlDoc.DocumentNode.SelectSingleNode("//*[@id=\"comment_total_normal\"]");
                        HtmlNode badCommentNode = htmlDoc.DocumentNode.SelectSingleNode("//*[@id=\"comment_total_bad\"]");

                        totalCommentCount = int.Parse(totalCommentNode.InnerText.Trim());
                        hCommentCount = int.Parse(goodCommentNode.InnerText.Trim());
                        mCommentCount = int.Parse(normalCommentNode.InnerText.Trim());
                        lCommentCount = int.Parse(badCommentNode.InnerText.Trim());

                        hPer = totalCommentCount == 0 ? null : (Nullable<decimal>)(((decimal)hCommentCount) / (decimal)totalCommentCount);
                        */

                        HtmlNodeCollection pNodes = htmlDoc.DocumentNode.SelectNodes("//*[@class=\"comment clearfix\"]/div");
                        if (pNodes != null)
                        {
                            foreach (HtmlNode pNode in pNodes)
                            {
                                HtmlNode pTitleNode = pNode.SelectSingleNode("./h5");
                                string pTitle = pTitleNode.InnerText.Replace(" ", "").Trim();
                                switch (pTitle)
                                {
                                    case "备注":
                                        {
                                            HtmlNode vNode = pNode.SelectSingleNode("./span");
                                            note = vNode.InnerText.Trim();
                                        }
                                        break;
                                    case "甜度":
                                        {
                                            HtmlNode imgNode = pNode.SelectSingleNode("./span/img");
                                            if (imgNode != null)
                                            {
                                                string imgUrl = imgNode.Attributes["src"].Value;
                                                int startIndex = imgUrl.LastIndexOf("-") + 1;
                                                int endIndex = imgUrl.LastIndexOf(".");
                                                if (endIndex >= 0 && startIndex >= 0 && endIndex > startIndex)
                                                {
                                                    sweetness = int.Parse(imgUrl.Substring(startIndex, endIndex - startIndex));
                                                }
                                            }
                                        }
                                        break;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        this.RunPage.InvokeAppendLogText("读取出错. url = " + pageUrl + ". " + ex.Message, LogLevelType.Error, true);
                        throw ex;
                    }
                    finally
                    {
                        if (commentTr != null)
                        {
                            commentTr.Close();
                            commentTr.Dispose();
                        }
                        if (htmlTr != null)
                        {
                            htmlTr.Close();
                            htmlTr.Dispose();
                        }
                    }

                    Dictionary<string, object> f2vs = new Dictionary<string, object>();
                    f2vs.Add("商品名称", productName);
                    if (!CommonUtil.IsNullOrBlank(productCurrentPrice))
                    {
                        f2vs.Add("价格", decimal.Parse(productCurrentPrice));
                    }
                    f2vs.Add("分类", categoryName);
                    f2vs.Add("规格", standard);
                    f2vs.Add("评论数", totalCommentCount);
                    f2vs.Add("好评", hCommentCount);
                    f2vs.Add("中评", mCommentCount);
                    f2vs.Add("差评", lCommentCount);
                    if (hPer != null)
                    {
                        f2vs.Add("好评度", hPer);
                    }
                    f2vs.Add("url", pageUrl);
                    if (!CommonUtil.IsNullOrBlank(productOldPrice))
                    {
                        f2vs.Add("参考价", decimal.Parse(productOldPrice));
                    }
                    f2vs.Add("商品编码", productCode);
                    f2vs.Add("分类编码", categoryCode);
                    f2vs.Add("备注", note);
                    f2vs.Add("甜度", sweetness);
                    f2vs.Add("城市", city);
                    resultEW.AddRow(f2vs);
                }
            }
        }
        #endregion
    }
}