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
using NetDataAccess.Base.Reader;
using NetDataAccess.Extended.Taobao.Common;

namespace NetDataAccess.Extended.Taobao.List
{
    /// <summary>
    /// GetAllListPages
    /// </summary>
    [System.Runtime.InteropServices.ComVisibleAttribute(true)]
    public class GetAllListPages : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            return this.GetShopInfos(listSheet);
        }

        private string GetShopUrl(string url)
        {
            int qIndex = url.IndexOf("?");
            if (qIndex >= 0)
            {
                return url.Substring(0, qIndex).Replace("/", "");
            }
            else
            {
                return url.Replace("/", "");
            }
        }

        private string GetShopId(string shopUrl)
        {
            int dIndex = shopUrl.IndexOf(".");
            return "https://" + shopUrl.Substring(0, dIndex);
        }

        private string GetShopType(string shopUrl)
        {
            int ldIndex = shopUrl.LastIndexOf(".");
            string tempStr = shopUrl.Substring(0, ldIndex);
            int ddIndex = tempStr.LastIndexOf(".");
            if (ddIndex >= 0)
            {
                return tempStr.Substring(ddIndex + 1).ToLower();
            }
            else
            {
                return tempStr.ToLower();
            }
        }

        private string GetShopProductListPageUrl(string shopId)
        {
            string url = "https://" + shopId + ".taobao.com/search.htm";
            return url;
        }

        /// <summary>
        /// 获取列表页里的店铺信息
        /// </summary>
        /// <param name="listSheet"></param>
        /// <returns></returns>
        private bool GetShopInfos(IListSheet listSheet)
        {
            bool succeed = true;
            string exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> columnDic = CommonUtil.InitStringIndexDic(new string[]{
                    "detailPageUrl",
                    "detailPageName", 
                    "cookie",
                    "grabStatus", 
                    "giveUpGrab",
                    "店铺名",
                    "店铺网址",
                    "店铺类型",
                    "淘宝店铺级别",
                    "卖家旺旺账号",
                    "所在地区",
                    "主营",
                    "SKU个数",
                    "描述相符",
                    "服务态度",
                    "物流服务",
                    "好评率",
                    "店铺关键字",
                    "关键字搜索页码",
                    "店铺号",
                    "所属行业"});
            string shopFirstPageUrlFilePath = Path.Combine(exportDir, "淘宝App店铺.xlsx");
            ExcelWriter ew = new ExcelWriter(shopFirstPageUrlFilePath, "List", columnDic, null);

            Dictionary<string, string> keywordShopIdToNull = new Dictionary<string, string>();
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"];
                string keyword = row["keyword"]; 
                string pageNum = row["pageNum"]; 
                string localFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir);
                TextReader tr = null;

                try
                {
                    tr = new StreamReader(localFilePath, Encoding.GetEncoding(((Proj_Detail_SingleLine)this.RunPage.Project.DetailGrabInfoObject).Encoding));
                    string webPageHtml = tr.ReadToEnd();

                    HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                    htmlDoc.LoadHtml(webPageHtml);

                    HtmlNodeCollection listNodeListA = htmlDoc.DocumentNode.SelectNodes("//li[@class=\"list-item\"]/ul/li[contains(@class,\"list-info\")]");
                    if (listNodeListA.Count > 0 )
                    {
                        this.GetShopItem(listNodeListA, keyword, pageNum, keywordShopIdToNull, ew); 
                    }

                    //这是根据滚动，动态构造的节点
                    HtmlNodeCollection listNodeListB = htmlDoc.DocumentNode.SelectNodes("//li[@class=\"list-item\"]/div/ul/li[contains(@class,\"list-info\")]");
                    if (listNodeListB.Count > 0)
                    { 
                        this.GetShopItem(listNodeListB, keyword, pageNum, keywordShopIdToNull, ew);
                    }
                }
                catch (Exception ex)
                {
                    if (tr != null)
                    {
                        tr.Dispose();
                        tr = null;
                    }
                    this.RunPage.InvokeAppendLogText("读取出错. " + ex.Message + " LocalPath = " + localFilePath, LogLevelType.Error, true);
                }
            }
            ew.SaveToDisk();
            return succeed;
        }

        private void GetShopItem(HtmlNodeCollection listNodeList, string keyword, string pageNum, Dictionary<string, string> keywordShopIdToNull, ExcelWriter ew)
        {
            for (int j = 0; j < listNodeList.Count; j++)
            {
                HtmlNode listNode = listNodeList[j];

                string shopName = "";
                string shopUrl = "";
                string shopId = "";
                string shopType = "";
                string shopLevel = "";
                string shopProductPageUrl = "";
                string wangwang = "";
                string district = "";
                string mainCat = "";
                string skuCount = "";
                string matchMark = "";
                string serviceMark = "";
                string transportMark = "";
                string goodPerc = "";
                string industryType = "";
                string shopNum = "";


                HtmlNode shopNameNode = listNode.SelectSingleNode("./h4/a[@class=\"shop-name J_shop_name\"]");
                string shopLinkUrl = shopNameNode.GetAttributeValue("href", "");

                //店铺地址
                shopUrl = this.GetShopUrl(shopLinkUrl);

                //店铺Id
                shopId = this.GetShopId(shopUrl); 

                //店铺名称
                shopName = CommonUtil.HtmlDecode(shopNameNode.InnerText.Trim());

                //店铺商品列表首页
                shopProductPageUrl = this.GetShopProductListPageUrl(shopId);


                //店铺类型，天猫or淘宝or飞猪
                HtmlNode tmallNode = listNode.SelectSingleNode("./h4/a[@class=\"icon-service-tianmao-large\"]");
                if (tmallNode != null)
                {
                    shopType = "天猫";
                }
                else
                {
                    HtmlNode feizhuNode = listNode.SelectSingleNode("./h4/a[@class=\"icon-fest-feizhudianpu\"]");
                    if (feizhuNode != null)
                    {
                        shopType = "飞猪";
                    }
                    else
                    {
                        shopType = "淘宝";
                    }
                }

                //店铺级别
                HtmlNode shopLevelClassNode = listNode.SelectSingleNode("./h4/a[contains(@class,\"rank\")]");
                if (shopLevelClassNode != null)
                {
                    string shopLevelClass = shopLevelClassNode.GetAttributeValue("class", "").Trim();
                    int lIndex = shopLevelClass.LastIndexOf("-");
                    if (lIndex >= 0)
                    {
                        shopLevel = shopLevelClass.Substring(lIndex + 1);
                    }
                }

                // "卖家旺旺账号"
                HtmlNode wangwangNode = listNode.SelectSingleNode("./p[@class=\"shop-info\"]/span[@class=\"shop-info-list\"]/a");
                if (wangwangNode != null)
                {
                    wangwang = wangwangNode.InnerText.Trim();
                }

                //"所在地区"
                HtmlNode districtNode = listNode.SelectSingleNode("./p[@class=\"shop-info\"]/span[@class=\"shop-address\"]");
                if (districtNode != null)
                {
                    district = districtNode.InnerText.Trim();
                }

                //"主营"
                HtmlNode mainCatNode = listNode.SelectSingleNode("./p[@class=\"main-cat\"]/a");
                if (mainCatNode != null)
                {
                    mainCat = CommonUtil.HtmlDecode(mainCatNode.InnerText.Trim());
                }

                //"SKU个数"
                HtmlNode skuCountNode = listNode.SelectSingleNode("./span[@class=\"pro-sale-num\"]/span[@class=\"info-sum\"]/em");
                if (skuCountNode != null)
                {
                    skuCount = skuCountNode.InnerText.Trim();
                }

                //"描述相符"、"服务态度"、"物流服务",
                //data-dsr="{"srn":"2262147","sgr":"99.05%","ind":"服饰鞋包","mas":"4.69","mg":"0.00%","sas":"4.75","sg":"0.00%","cas":"4.73","cg":"0.00%","encryptedUserId":"UMGNLMCHSOmNG"}"
                HtmlNode markNode = listNode.SelectSingleNode("./div[@class=\"valuation clearfix\"]/div[@class=\"descr J_descr target-hint-descr\"]");
                if (markNode != null)
                {
                    string markJson = markNode.GetAttributeValue("data-dsr", "");
                    JObject rootJo = JObject.Parse(markJson);
                    shopNum = rootJo["srn"].ToString();
                    industryType = rootJo["ind"].ToString();
                    matchMark = rootJo["mas"].ToString();
                    serviceMark = rootJo["sas"].ToString();
                    transportMark = rootJo["cas"].ToString();
                    goodPerc = rootJo["sgr"].ToString(); 
                } 

                Dictionary<string, object> shopInfo = new Dictionary<string, object>();
                string keywordShopId = keyword + "_" + shopId;
                if (!keywordShopIdToNull.ContainsKey(keywordShopId))
                {
                    keywordShopIdToNull.Add(keywordShopId, null);
                    shopInfo.Add("detailPageUrl", shopProductPageUrl);
                    shopInfo.Add("detailPageName", keywordShopId);
                    shopInfo.Add("店铺名", shopName);
                    shopInfo.Add("店铺网址", shopUrl);
                    shopInfo.Add("店铺类型", shopType);
                    shopInfo.Add("淘宝店铺级别", shopLevel);
                    shopInfo.Add("卖家旺旺账号", wangwang);
                    shopInfo.Add("所在地区", district);
                    shopInfo.Add("主营", mainCat);
                    shopInfo.Add("SKU个数", skuCount);
                    shopInfo.Add("描述相符", matchMark);
                    shopInfo.Add("服务态度", serviceMark);
                    shopInfo.Add("物流服务", transportMark);
                    shopInfo.Add("好评率", goodPerc);
                    shopInfo.Add("店铺关键字", keyword);
                    shopInfo.Add("关键字搜索页码", pageNum);
                    shopInfo.Add("所属行业", industryType);
                    shopInfo.Add("店铺号", shopNum);
                    ew.AddRow(shopInfo);
                }
            }
        }

        public override void WebBrowserHtml_AfterPageLoaded(string pageUrl, Dictionary<string, string> listRow, WebBrowser webBrowser)
        {
            //滚动到页面最下面
            ProcessWebBrowser.AutoScroll(this.RunPage, webBrowser, 0, 5000, 1000, 1000, 1000);
        }
        public override void CheckRequestCompleteFile(string webPageText, Dictionary<string, string> listRow)
        {
            if (!webPageText.Contains("下一页"))
            {
                throw new Exception("Uncompleted webpage request.");
            } 
        }
    }
}