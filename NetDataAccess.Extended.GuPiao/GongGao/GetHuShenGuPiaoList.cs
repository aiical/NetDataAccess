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

namespace NetDataAccess.Extended.GuPiao
{
    public class GetHuShenGuPiaoList : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            this.GetGongGaoListFirstPages(listSheet);
            this.GetGuPiaoList(listSheet);
            return true;
        }

        private string GetStockExchange(string code)
        {
            if (code.StartsWith("60") || code.StartsWith("900"))
            {
                return "沪市";
            }
            else
            {
                return "深市";
            }
        }

        private string GetCategory(string code)
        {
            if (code.StartsWith("60") || code.StartsWith("000"))
            {
                return "A股";
            }
            else if (code.StartsWith("900") || code.StartsWith("200"))
            {
                return "B股";
            }
            else if (code.StartsWith("002"))
            {
                return "中小板";
            }
            else if (code.StartsWith("300"))
            {
                return "创业板";
            }
            else
            {
                return "";
            }
        }



        private void GetGongGaoListFirstPages(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("pinyin", 5);
            resultColumnDic.Add("zwjc", 6);
            resultColumnDic.Add("code", 7);
            resultColumnDic.Add("orgId", 8);
            resultColumnDic.Add("stockExchange", 9);
            resultColumnDic.Add("category", 10);
            resultColumnDic.Add("formData", 11);
            string resultFilePath = Path.Combine(exportDir, "沪深股票公告列表页首页.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
             
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"];
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string localFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir);
                    TextReader tr = null;

                    try
                    {
                        tr = new StreamReader(localFilePath, Encoding.UTF8);
                        string js = tr.ReadToEnd();

                        JObject rootJo = JObject.Parse(js);
                        JArray gpListJsons = rootJo.SelectToken("stockList") as JArray;

                        for (int j = 0; j < gpListJsons.Count; j++)
                        {
                            JObject gpJson = gpListJsons[j] as JObject;
                            string orgId = gpJson.GetValue("orgId").ToString();
                            string pinyin = gpJson.GetValue("pinyin").ToString().ToUpper();
                            string code = gpJson.GetValue("code").ToString();
                            string zwjc = gpJson.GetValue("zwjc").ToString();
                            string category = this.GetCategory(code);
                            string stockExchange = this.GetStockExchange(code);

                            string nameEncodeString = CommonUtil.UrlEncode(zwjc);
                            Dictionary<string, string> f2vs = new Dictionary<string, string>();
                            f2vs.Add("detailPageUrl", "http://www.cninfo.com.cn/cninfo-new/announcement/query?_=" + code);
                            f2vs.Add("detailPageName", code);
                            f2vs.Add("cookie", "JSESSIONID=7AFC587425A130F0BE574853993CC056; _sp_id.2141=a2cd0eab-a356-499a-a840-965710788f1e.1516760763.1.1516760766.1516760763.5277b143-53fc-464a-94e3-6d5ddeb8844c; cninfo_search_record_cookie=600656|%E6%96%B0%E9%83%BD%E9%80%80|; JSESSIONID=D2D32C07D9EC7A73CC0D6549AA726F54");
                            f2vs.Add("zwjc", zwjc);
                            f2vs.Add("code", code);
                            f2vs.Add("pinyin", pinyin);
                            f2vs.Add("orgId", orgId);
                            f2vs.Add("category", category);
                            f2vs.Add("stockExchange", stockExchange);
                            f2vs.Add("formData", "stock=" + code + "&searchkey=&category=&pageNum=1&pageSize=30&column=szse_main&tabName=fulltext&sortName=&sortType=&limit=&seDate=");
                            resultEW.AddRow(f2vs);
                        }

                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                    finally
                    {
                        if (tr != null)
                        {
                            tr.Close();
                            tr.Dispose();
                        }
                    }
                }
            } 
            resultEW.SaveToDisk();
        }

        private void GetGuPiaoList(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("code", 0);
            resultColumnDic.Add("pinyin", 1);
            resultColumnDic.Add("zwjc", 2);
            resultColumnDic.Add("orgId", 3);
            resultColumnDic.Add("category", 4);
            resultColumnDic.Add("stockExchange", 5); 
            string resultFilePath = Path.Combine(exportDir, "沪深股票列表.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;
            List<Dictionary<string, string>> gupiaoList = new List<Dictionary<string, string>>(); 

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"];
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string localFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir);
                    TextReader tr = null;

                    try
                    {
                        tr = new StreamReader(localFilePath, Encoding.UTF8);
                        string js = tr.ReadToEnd();

                        JObject rootJo = JObject.Parse(js);
                        JArray gpListJsons = rootJo.SelectToken("stockList") as JArray;

                        for (int j = 0; j < gpListJsons.Count; j++)
                        {
                            JObject gpJson = gpListJsons[j] as JObject;
                            string orgId = gpJson.GetValue("orgId").ToString();
                            string pinyin = gpJson.GetValue("pinyin").ToString().ToUpper();
                            string code = gpJson.GetValue("code").ToString();
                            string zwjc = gpJson.GetValue("zwjc").ToString();
                            string category = this.GetCategory(code);
                            string stockExchange = this.GetStockExchange(code);

                            Dictionary<string, string> f2vs = new Dictionary<string, string>();
                            f2vs.Add("zwjc", zwjc);
                            f2vs.Add("code", code);
                            f2vs.Add("pinyin", pinyin);
                            f2vs.Add("orgId", orgId);
                            f2vs.Add("category", category);
                            f2vs.Add("stockExchange", stockExchange);
                            resultEW.AddRow(f2vs);
                        }

                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                    finally
                    {
                        if (tr != null)
                        {
                            tr.Close();
                            tr.Dispose();
                        }
                    }
                }
            }
            resultEW.SaveToDisk();
        }
    }
}