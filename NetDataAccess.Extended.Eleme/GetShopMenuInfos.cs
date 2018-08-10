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

namespace NetDataAccess.Extended.Eleme
{
    public class GetShopMenuInfos : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            string[] parameters = this.Parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            string dirPath = parameters[0];
            Dictionary<string, Dictionary<string, string>> allShopDic = this.GetAllShopDic(listSheet);

            for (int i = 1; i < parameters.Length; i++)
            {
                string city = parameters[i];
                ExcelReader cityEr = this.GetCityER(dirPath, city);

                this.GetCategoryMenuMaps(cityEr, city, allShopDic);
                this.GetMenus(cityEr, city, allShopDic);
            }
            return true;
        }


        private Dictionary<string, Dictionary<string,string>> GetAllShopDic(IListSheet listSheet)
        {
            int rowCount = listSheet.RowCount;
            Dictionary<string, Dictionary<string, string>> allShopDic = new Dictionary<string, Dictionary<string, string>>();
            for (int i = 0; i < rowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string url = row[SysConfig.DetailPageUrlFieldName];
                allShopDic.Add(url, row);
            }
            return allShopDic;
        }

        private ExcelReader GetCityER(string dirPath, string city)
        {
            string filePath = Path.Combine(dirPath, "饿了么_店铺详情页_" + city + ".xlsx");
            ExcelReader er = new ExcelReader(filePath);
            return er;
        }

        /*
        private Dictionary<string, string> GetCityShopDic(string dirPath, string city)
        {
            string filePath = Path.Combine(dirPath, "饿了么_店铺详情页_" + city + ".xlsx");
            ExcelReader er = new ExcelReader(filePath);
            int rowCount = er.GetRowCount();
            Dictionary<string, string> shopDic = new Dictionary<string, string>();
            for (int i = 0; i < rowCount; i++)
            {
                Dictionary<string,string> row = er.GetFieldValues(i);
                string url = row[SysConfig.DetailPageUrlFieldName];
                shopDic.Add(url, null);
            }
            return shopDic;
        }
         * */

        public override void WebRequestHtml_BeforeSendRequest(string pageUrl, Dictionary<string, string> listRow, Base.Web.NDAWebClient client)
        {
            base.WebRequestHtml_BeforeSendRequest(pageUrl, listRow, client);
            string xShard = listRow["xShard"];

            client.Headers.Add("x-shard", xShard);
            client.Headers.Add("content-type", "application/json");
        }
        public override byte[] GetRequestData_BeforeSendRequest(string pageUrl, Dictionary<string, string> listRow, Encoding encoding)
        {
            string dataText = listRow["dataText"];
            byte[] dataArray = Encoding.UTF8.GetBytes(dataText);
            return dataArray;
        }

        private void GetCategoryMenuMaps(ExcelReader cityEr, string city, Dictionary<string,Dictionary<string,string>> allShopDic)
        { 
            CsvWriter cw = this.CreateCategoryMenuMapsFileWriter(city); 
            string sourceDir = this.RunPage.GetDetailSourceFileDir();
            int rowCount = cityEr.GetRowCount();
            for (int i = 0; i < rowCount; i++)
            { 
                Dictionary<string, string> cityShopRow = cityEr.GetFieldValues(i);
                string detailPageUrl = cityShopRow[SysConfig.DetailPageUrlFieldName];
                if (allShopDic.ContainsKey(detailPageUrl))
                {
                    Dictionary<string, string> listRow = allShopDic[detailPageUrl];
                    bool giveUp = "Y".Equals(listRow[SysConfig.GiveUpGrabFieldName]);
                    if (!giveUp)
                    {

                        string filePath = this.RunPage.GetFilePath(detailPageUrl, sourceDir);

                        string jsonText = FileHelper.GetTextFromFile(filePath);
                        try
                        {
                            JObject rootJo = JObject.Parse(jsonText);
                            JObject menuJo = rootJo.GetValue("menu") as JObject;
                            bool gotMenu = false;
                            if (menuJo != null)
                            {
                                string bodyJo = menuJo.GetValue("body").ToString();
                                if (bodyJo != null && bodyJo.Length > 0)
                                {
                                    JArray categoryArray = JArray.Parse(bodyJo);
                                    if (categoryArray.Count > 0)
                                    {
                                        gotMenu = true;

                                        for (int j = 0; j < categoryArray.Count; j++)
                                        {
                                            JObject categoryJo = categoryArray[j] as JObject;
                                            JArray foodArray = categoryJo.GetValue("foods") as JArray;
                                            string categoryId = categoryJo.GetValue("id") == null ? "" : categoryJo.GetValue("id").ToString();
                                            string categoryName = categoryJo.GetValue("name") == null ? "" : categoryJo.GetValue("name").ToString();
                                            string categoryDescription = categoryJo.GetValue("description") == null ? "" : categoryJo.GetValue("description").ToString();
                                            if (foodArray != null)
                                            {
                                                for (int k = 0; k < foodArray.Count; k++)
                                                {
                                                    JObject foodJo = foodArray[k] as JObject;
                                                    string foodId = foodJo.GetValue("item_id") == null ? "" : foodJo.GetValue("item_id").ToString();
                                                    string foodName = foodJo.GetValue("name") == null ? "" : foodJo.GetValue("name").ToString();
                                                    string rating = foodJo.GetValue("rating") == null ? "" : foodJo.GetValue("rating").ToString();
                                                    string monthSales = foodJo.GetValue("month_sales") == null ? "" : foodJo.GetValue("month_sales").ToString();
                                                    string ratingCount = foodJo.GetValue("rating_count") == null ? "" : foodJo.GetValue("rating_count").ToString();
                                                    string statisfyCount = foodJo.GetValue("statisfy_count") == null ? "" : foodJo.GetValue("statisfy_count").ToString();
                                                    string statisfyRate = foodJo.GetValue("statisfy_rate") == null ? "" : foodJo.GetValue("statisfy_rate").ToString();
                                                    string minPurchase = foodJo.GetValue("min_purchase") == null ? "" : foodJo.GetValue("min_purchase").ToString();

                                                    Dictionary<string, string> categoryFoodRow = new Dictionary<string, string>();
                                                    categoryFoodRow.Add("id", listRow["id"]);
                                                    categoryFoodRow.Add("name", listRow["name"]);
                                                    categoryFoodRow.Add("address", listRow["address"]);
                                                    categoryFoodRow.Add("description", listRow["description"]);
                                                    categoryFoodRow.Add("latitude", listRow["latitude"]);
                                                    categoryFoodRow.Add("longitude", listRow["longitude"]);
                                                    categoryFoodRow.Add("phone", listRow["phone"]);
                                                    categoryFoodRow.Add("promotion_info", listRow["promotion_info"]);

                                                    categoryFoodRow.Add("categoryId", categoryId);
                                                    categoryFoodRow.Add("categoryName", categoryName);
                                                    categoryFoodRow.Add("categoryDescription", categoryDescription);
                                                    categoryFoodRow.Add("foodId", foodId);
                                                    categoryFoodRow.Add("foodName", foodName);
                                                    categoryFoodRow.Add("rating", rating);
                                                    categoryFoodRow.Add("monthSales", monthSales);
                                                    categoryFoodRow.Add("ratingCount", ratingCount);
                                                    categoryFoodRow.Add("statisfyCount", statisfyCount);
                                                    categoryFoodRow.Add("statisfyRate", statisfyRate);
                                                    categoryFoodRow.Add("minPurchase", minPurchase);
                                                    cw.AddRow(categoryFoodRow);

                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            /*
                            if (!gotMenu)
                            {
                                this.RunPage.InvokeAppendLogText("(" + (i + 1).ToString() + "/" + rowCount.ToString() + ")删除文件 " + filePath, LogLevelType.System, true);
                                File.Delete(filePath);
                            }*/
                        }
                        catch (Exception ex)
                        {
                            this.RunPage.InvokeAppendLogText(ex.Message + ". FilePath = " + filePath, LogLevelType.System, true);

                        }
                    }
                }
            }
            cw.SaveToDisk();
        }

        public override void CheckRequestCompleteFile(string webPageText, Dictionary<string, string> listRow)
        {/*
            string detailPageUrl = listRow[SysConfig.DetailPageUrlFieldName];
            JObject rootJo = JObject.Parse(webPageText);
            JObject menuJo = rootJo.GetValue("menu") as JObject;
            bool gotMenu = false;
            if (menuJo != null)
            {
                string bodyJo = menuJo.GetValue("body").ToString();
                if (bodyJo != null && bodyJo.Length > 0)
                {
                    JArray categoryArray = JArray.Parse(bodyJo);
                    if (categoryArray.Count > 0)
                    {
                        gotMenu = true;
                    }
                }
            }
            if (!gotMenu)
            {
                throw new Exception("菜单为空! Url = " + detailPageUrl);
            }*/
        }

        private CsvWriter CreateCategoryMenuMapsFileWriter(string city)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            string resultFilePath = Path.Combine(exportDir, "饿了么_店铺菜单与分类对照_" + city + ".csv");
            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("id", 0);
            resultColumnDic.Add("name", 1);
            resultColumnDic.Add("address", 2);
            resultColumnDic.Add("description", 3);
            resultColumnDic.Add("latitude", 4);
            resultColumnDic.Add("longitude", 5);
            resultColumnDic.Add("phone", 6);
            resultColumnDic.Add("promotion_info", 7); 

            resultColumnDic.Add("categoryId", 8);
            resultColumnDic.Add("categoryName", 9);
            resultColumnDic.Add("categoryDescription", 10);
            resultColumnDic.Add("foodId", 11);
            resultColumnDic.Add("foodName", 12);
            resultColumnDic.Add("rating", 13);
            resultColumnDic.Add("monthSales", 14);
            resultColumnDic.Add("ratingCount", 15);
            resultColumnDic.Add("statisfyCount", 16);
            resultColumnDic.Add("statisfyRate", 17);
            resultColumnDic.Add("minPurchase", 18);

            CsvWriter resultEW = new CsvWriter(resultFilePath, resultColumnDic);
            return resultEW;
        }


        private void GetMenus(ExcelReader cityEr, string city, Dictionary<string, Dictionary<string, string>> allShopDic)
        {
            Dictionary<string, string> menuIdDic = new Dictionary<string, string>();
            CsvWriter cw = this.CreateMenuFileWriter(city);
            string sourceDir = this.RunPage.GetDetailSourceFileDir();
            int rowCount = cityEr.GetRowCount();
            for (int i = 0; i < rowCount; i++)
            { 
                Dictionary<string, string> cityShopRow = cityEr.GetFieldValues(i);
                string detailPageUrl = cityShopRow[SysConfig.DetailPageUrlFieldName];
                if (allShopDic.ContainsKey(detailPageUrl))
                {
                    Dictionary<string, string> listRow = allShopDic[detailPageUrl];
                    bool giveUp = "Y".Equals(listRow[SysConfig.GiveUpGrabFieldName]);
                    if (!giveUp)
                    { 
                        string filePath = this.RunPage.GetFilePath(detailPageUrl, sourceDir);

                        string jsonText = FileHelper.GetTextFromFile(filePath);
                        try
                        {
                            JObject rootJo = JObject.Parse(jsonText);
                            JObject menuJo = rootJo.GetValue("menu") as JObject;
                            bool gotMenu = false;
                            if (menuJo != null)
                            {
                                string bodyJo = menuJo.GetValue("body").ToString();
                                if (bodyJo != null && bodyJo.Length > 0)
                                {
                                    JArray categoryArray = JArray.Parse(bodyJo);
                                    if (categoryArray.Count > 0)
                                    {
                                        gotMenu = true;

                                        for (int j = 0; j < categoryArray.Count; j++)
                                        {
                                            JObject categoryJo = categoryArray[j] as JObject;
                                            JArray foodArray = categoryJo.GetValue("foods") as JArray;
                                            string categoryId = categoryJo.GetValue("id") == null ? "" : categoryJo.GetValue("id").ToString();
                                            string categoryName = categoryJo.GetValue("name") == null ? "" : categoryJo.GetValue("name").ToString();
                                            string categoryDescription = categoryJo.GetValue("description") == null ? "" : categoryJo.GetValue("description").ToString();
                                            if (foodArray != null)
                                            {
                                                for (int k = 0; k < foodArray.Count; k++)
                                                {
                                                    JObject foodJo = foodArray[k] as JObject;
                                                    string foodId = foodJo.GetValue("item_id") == null ? "" : foodJo.GetValue("item_id").ToString();

                                                    string menuFullId = listRow["id"] + "_" + foodId;
                                                    if (!menuIdDic.ContainsKey(menuFullId))
                                                    {
                                                        menuIdDic.Add(menuFullId, null);

                                                        string foodName = foodJo.GetValue("name") == null ? "" : foodJo.GetValue("name").ToString();
                                                        string rating = foodJo.GetValue("rating") == null ? "" : foodJo.GetValue("rating").ToString();
                                                        string monthSales = foodJo.GetValue("month_sales") == null ? "" : foodJo.GetValue("month_sales").ToString();
                                                        string ratingCount = foodJo.GetValue("rating_count") == null ? "" : foodJo.GetValue("rating_count").ToString();
                                                        string statisfyCount = foodJo.GetValue("statisfy_count") == null ? "" : foodJo.GetValue("statisfy_count").ToString();
                                                        string statisfyRate = foodJo.GetValue("statisfy_rate") == null ? "" : foodJo.GetValue("statisfy_rate").ToString();
                                                        string minPurchase = foodJo.GetValue("min_purchase") == null ? "" : foodJo.GetValue("min_purchase").ToString();

                                                        Dictionary<string, string> categoryFoodRow = new Dictionary<string, string>();
                                                        categoryFoodRow.Add("id", listRow["id"]);
                                                        categoryFoodRow.Add("name", listRow["name"]);
                                                        categoryFoodRow.Add("address", listRow["address"]);
                                                        categoryFoodRow.Add("description", listRow["description"]);
                                                        categoryFoodRow.Add("latitude", listRow["latitude"]);
                                                        categoryFoodRow.Add("longitude", listRow["longitude"]);
                                                        categoryFoodRow.Add("phone", listRow["phone"]);
                                                        categoryFoodRow.Add("promotion_info", listRow["promotion_info"]);

                                                        categoryFoodRow.Add("foodId", foodId);
                                                        categoryFoodRow.Add("foodName", foodName);
                                                        categoryFoodRow.Add("rating", rating);
                                                        categoryFoodRow.Add("monthSales", monthSales);
                                                        categoryFoodRow.Add("ratingCount", ratingCount);
                                                        categoryFoodRow.Add("statisfyCount", statisfyCount);
                                                        categoryFoodRow.Add("statisfyRate", statisfyRate);
                                                        categoryFoodRow.Add("minPurchase", minPurchase);
                                                        cw.AddRow(categoryFoodRow);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            /*
                            if (!gotMenu)
                            {
                                this.RunPage.InvokeAppendLogText("(" + (i + 1).ToString() + "/" + rowCount.ToString() + ")删除文件 " + filePath, LogLevelType.System, true);
                                File.Delete(filePath);
                            }*/
                        }
                        catch (Exception ex)
                        {
                            this.RunPage.InvokeAppendLogText(ex.Message + ". FilePath = " + filePath, LogLevelType.System, true);

                        }
                    }
                }
            }
            cw.SaveToDisk();
        }


        private CsvWriter CreateMenuFileWriter(string city)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            string resultFilePath = Path.Combine(exportDir, "饿了么_店铺菜单_" + city + ".csv");
            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("id", 0);
            resultColumnDic.Add("name", 1);
            resultColumnDic.Add("address", 2);
            resultColumnDic.Add("description", 3);
            resultColumnDic.Add("latitude", 4);
            resultColumnDic.Add("longitude", 5);
            resultColumnDic.Add("phone", 6);
            resultColumnDic.Add("promotion_info", 7); 

            resultColumnDic.Add("foodId", 8);
            resultColumnDic.Add("foodName", 9);
            resultColumnDic.Add("rating", 10);
            resultColumnDic.Add("monthSales", 11);
            resultColumnDic.Add("ratingCount", 12);
            resultColumnDic.Add("statisfyCount", 13);
            resultColumnDic.Add("statisfyRate", 14);
            resultColumnDic.Add("minPurchase", 15);

            CsvWriter resultEW = new CsvWriter(resultFilePath, resultColumnDic);
            return resultEW;
        }
    }
}