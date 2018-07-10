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
    public class WomaiDetailPageInfoJG : CustomProgramBase
    {
        public bool Run(string parameters, IListSheet listSheet)
        {
            return this.GenerateDetailPageInfo(listSheet);
        }

        private bool GenerateDetailPageInfo(IListSheet listSheet)
        {
            bool succeed = true;
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                "商品名称",
                "商品编码",
                "一级分类编码", 
                "一级分类", 
                "二级分类编码", 
                "二级分类", 
                "抢购价",
                "VIP价",
                "规格", 
                "品牌", 
                "净含量", 
                "产品毛重", 
                "产地", 
                "保质期", 
                "地区",
                "营销栏目",
                "商品页Url",
                "priceUrl",
                "priceLocalPath"});


            Dictionary<string, string> resultColumnFormat = new Dictionary<string, string>();
            resultColumnFormat.Add("抢购价", "#,##0.00"); 
            resultColumnFormat.Add("VIP价", "#,##0.00");

            string readDetailDir = this.RunPage.GetReadFileDir();
            string exportDir = this.RunPage.GetExportDir();
            string resultFilePath = Path.Combine(exportDir, "我买网商品详情" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, resultColumnFormat);

            GetList(listSheet, pageSourceDir, resultEW); 

            resultEW.SaveToDisk(); 

            return succeed;
        }

        private void GetList(IListSheet listSheet, string pageSourceDir, ExcelWriter resultEW)
        {
            try
            { 
                for (int i = 0; i < listSheet.RowCount; i++)
                {
                    Dictionary<string, string> row = listSheet.GetRow(i);
                    string pageUrl = listSheet.PageUrlList[i];
                    string pageName = listSheet.PageNameList[i];
                    string localFilePath = this.RunPage.GetFilePath(pageUrl, pageSourceDir);
                    Nullable<decimal> qianggoujia = null;
                    Nullable<decimal> vipjia = null;
                    string productSysNo = row["productSysNo"];
                    string name = row["商品名称"];
                    string category1Code = row["category1Code"];
                    string category2Code = row["category2Code"];
                    string category1Name = row["category1Name"];
                    string category2Name = row["category2Name"];
                    string district = row["district"];
                    string guige = row["规格"];
                    string pinpai = row["品牌"];
                    string jinghanliang = row["净含量"];
                    string chanpinmaozhong = row["产品毛重"];
                    string chandi = row["产地"];
                    string baozhiqi = row["保质期"];
                    string yingxiaolanmu = row["营销栏目"];
                    string productUrl = row["商品页URL"]; 

                    TextReader tr = null;

                    try
                    {
                        tr = new StreamReader(localFilePath);
                        string webPageJson = tr.ReadToEnd();
                        int beginIndex = webPageJson.IndexOf("(") + 1;
                        int jsonLength = webPageJson.LastIndexOf(")") - beginIndex;
                        webPageJson = webPageJson.Substring(beginIndex, jsonLength);
                        JObject rootJo = JObject.Parse(webPageJson);
                        JArray resultArray = (JArray)rootJo["result"];
                        if (resultArray.Count > 0)
                        {
                            JObject priceJo = (JObject)((JObject)(resultArray)[0])["price"];
                            if (priceJo["buyPrice"] != null)
                            {
                                qianggoujia = decimal.Parse(priceJo["buyPrice"]["priceValue"].ToString());
                            }
                            if (priceJo["VIPPrice"] != null)
                            {
                                vipjia = decimal.Parse(priceJo["VIPPrice"]["priceValue"].ToString());
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        this.RunPage.InvokeAppendLogText("读取出错. url = " + pageUrl + ". " + ex.Message, LogLevelType.Error, true);
                        throw ex;
                    }

                    Dictionary<string, object> f2vs = new Dictionary<string, object>(); 
                    f2vs.Add("一级分类编码", category1Code);
                    f2vs.Add("二级分类编码", category2Code);
                    f2vs.Add("一级分类", category1Name);
                    f2vs.Add("二级分类", category2Name);
                    f2vs.Add("地区", district);
                    f2vs.Add("商品编码", productSysNo);
                    f2vs.Add("抢购价", qianggoujia);
                    f2vs.Add("VIP价", vipjia);
                    f2vs.Add("商品名称", name);
                    f2vs.Add("品牌", pinpai);
                    f2vs.Add("净含量", jinghanliang);
                    f2vs.Add("产品毛重", chanpinmaozhong);
                    f2vs.Add("产地", chandi);
                    f2vs.Add("保质期", baozhiqi);
                    f2vs.Add("规格", guige);
                    f2vs.Add("营销栏目", yingxiaolanmu);
                    f2vs.Add("商品页Url", productUrl);
                    f2vs.Add("priceUrl", pageUrl);
                    f2vs.Add("priceLocalPath", localFilePath); 
                    resultEW.AddRow(f2vs);
                }
            }
            catch (Exception ex)
            { 
                throw ex;
            }
        }
    }
}