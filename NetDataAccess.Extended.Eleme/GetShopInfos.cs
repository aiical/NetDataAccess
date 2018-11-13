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
    public class GetShopInfos : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            //this.GetCityToShops(listSheet);
            this.GetShopPageUrls(listSheet);
            return true;
        }

        public override void GetDataByOtherAccessType(Dictionary<string, string> listRow)
        {
            string detailPageUrl = listRow[SysConfig.DetailPageUrlFieldName];
            string urlFormat = listRow["urlFormat"];
            string lat = listRow["lat"];
            string lng = listRow["lng"];
            string sourceDir = this.RunPage.GetDetailSourceFileDir();
            string pointShopsFilePath = this.RunPage.GetFilePath(detailPageUrl, sourceDir);
            if (!File.Exists(pointShopsFilePath))
            {
                string pointShopDir = Path.Combine(sourceDir, detailPageUrl+"_dir");
                if (!Directory.Exists(pointShopDir))
                {
                    Directory.CreateDirectory(pointShopDir);
                }

                int pageCount = 0;
                bool needGetNextPage = true;
                while (needGetNextPage)
                {
                    int pageIndex = pageCount;
                    string nextListPageUrl = this.GetNextListPageUrl(urlFormat, lat, lng, pageIndex);
                    string localPath = this.RunPage.GetFilePath(nextListPageUrl, pointShopDir);
                    if (!File.Exists(localPath))
                    {
                        string pageText = this.RunPage.GetTextByRequest(nextListPageUrl, listRow, false, 0, 30000, Encoding.UTF8, null, null, false, Proj_DataAccessType.WebRequestHtml, null, 0);

                        JObject rootJo = JObject.Parse(pageText);
                        JArray itemArray = rootJo.GetValue("items") as JArray;
                        if (itemArray.Count == 0)
                        {
                            needGetNextPage = false;
                        }
                        else
                        {
                            pageCount++;
                            CommonUtil.CreateFileDirectory(localPath);
                            FileHelper.SaveTextToFile(pageText, localPath);
                        } 
                    }
                    else
                    {
                        pageCount++;
                    }
                }

                this.SaveShopsToPointFile(pointShopsFilePath, detailPageUrl, pageCount, pointShopDir, urlFormat,  lat, lng);
            }
        }

        private void SaveShopsToPointFile(string subCategoryFilePath, string detailPageUrl, int pageCount, string pointShopDir, string urlFormat, string lat, string lng)
        {
            ExcelWriter pointShopsEW = this.CreatePointShopsWriter(subCategoryFilePath);
            Dictionary<string, string> urlDic = new Dictionary<string, string>();
            for (int i = 0; i < pageCount; i++)
            {
                int pageIndex = i;
                string nextListPageUrl = this.GetNextListPageUrl(urlFormat, lat, lng, pageIndex);
                string localPath = this.RunPage.GetFilePath(nextListPageUrl, pointShopDir);
                string pageText = FileHelper.GetTextFromFile(localPath);

                JObject rootJo = JObject.Parse(pageText);
                JArray itemArray = rootJo.GetValue("items") as JArray;

                for (int j = 0; j < itemArray.Count; j++)
                {
                    try
                    {
                        JObject itemJo = (itemArray[j] as JObject).GetValue("restaurant") as JObject;
                        if (itemJo != null)
                        {
                            string address = itemJo.GetValue("address").ToString();
                            string description = itemJo.GetValue("description").ToString();
                            string id = itemJo.GetValue("id").ToString();
                            string latitude = itemJo.GetValue("latitude").ToString();
                            string longitude = itemJo.GetValue("longitude").ToString();
                            string name = itemJo.GetValue("name").ToString();
                            string phone = itemJo.GetValue("phone") == null ? "" : itemJo.GetValue("phone").ToString();
                            string promotion_info = itemJo.GetValue("promotion_info") == null ? "" : itemJo.GetValue("promotion_info").ToString();


                            if (!urlDic.ContainsKey(id))
                            {
                                urlDic.Add(id, null);

                                Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                f2vs.Add("address", address);
                                f2vs.Add("description", description);
                                f2vs.Add("id", id);
                                f2vs.Add("latitude", latitude);
                                f2vs.Add("longitude", longitude);
                                f2vs.Add("name", name);
                                f2vs.Add("phone", phone);
                                f2vs.Add("promotion_info", promotion_info);

                                pointShopsEW.AddRow(f2vs);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
            }
            pointShopsEW.SaveToDisk();
        }

        private string GetNextListPageUrl(string urlFormat, string lat, string lng, int pageIndex)
        {
            string url = urlFormat.Replace("##lat##", lat).Replace("##lng##", lng).Replace("##offset##", (pageIndex * 8).ToString());
            return url;
        }

        private ExcelWriter CreatePointShopsWriter(string subCategoryFilePath)
        {
            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();

            resultColumnDic.Add("address", 0);
            resultColumnDic.Add("description", 1);
            resultColumnDic.Add("id", 2);
            resultColumnDic.Add("latitude", 3);
            resultColumnDic.Add("longitude", 4);
            resultColumnDic.Add("name", 5);
            resultColumnDic.Add("phone", 6);
            resultColumnDic.Add("promotion_info", 7);

            ExcelWriter resultEW = new ExcelWriter(subCategoryFilePath, "List", resultColumnDic, null);
            return resultEW;
        }

        private CsvWriter CreateCityToShopsWriter(string mapFilePath)
        {
            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("address", 0);
            resultColumnDic.Add("description", 1);
            resultColumnDic.Add("id", 2);
            resultColumnDic.Add("latitude", 3);
            resultColumnDic.Add("longitude", 4);
            resultColumnDic.Add("name", 5);
            resultColumnDic.Add("phone", 6);
            resultColumnDic.Add("promotion_info", 7);
            resultColumnDic.Add("searchLat", 8);
            resultColumnDic.Add("searchLng", 9);
            resultColumnDic.Add("elemeCity", 10);
            CsvWriter resultEW = new CsvWriter(mapFilePath, resultColumnDic);
            return resultEW;
        }


        private void GetCityToShops(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            string resultFilePath = Path.Combine(exportDir, "饿了么_城市与店铺对照.csv");

            CsvWriter resultEW = this.CreateCityToShopsWriter(resultFilePath);

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailPageUrl = row[SysConfig.DetailPageUrlFieldName];
                string elemeCity = row["elemeCity"];
                string searchLat = row["lat"];
                string searchLng = row["lng"];

                string sourceDir = this.RunPage.GetDetailSourceFileDir();
                string shopsFilePath = this.RunPage.GetFilePath(detailPageUrl, sourceDir);

                ExcelReader er = new ExcelReader(shopsFilePath);
                int rowCount = er.GetRowCount();
                for (int j = 0; j < rowCount; j++)
                {
                    Dictionary<string, string> subRow = er.GetFieldValues(j);

                    Dictionary<string, string> mapRow = new Dictionary<string, string>();
                    mapRow.Add("address", subRow["address"]);
                    mapRow.Add("description", subRow["description"]);
                    mapRow.Add("id", subRow["id"]);
                    mapRow.Add("latitude", subRow["latitude"]);
                    mapRow.Add("longitude", subRow["longitude"]);
                    mapRow.Add("name", subRow["name"]);
                    mapRow.Add("phone", subRow["phone"]);
                    mapRow.Add("promotion_info", subRow["promotion_info"]);
                    mapRow.Add("searchLat", searchLat);
                    mapRow.Add("searchLng", searchLng);
                    mapRow.Add("elemeCity", elemeCity);
                    resultEW.AddRow(mapRow);
                }
            }
            resultEW.SaveToDisk();
        }

        private ExcelWriter CreateDetailFileWriter(string city)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            string resultFilePath = Path.Combine(exportDir, "饿了么_店铺详情页_" + city + ".xlsx");
            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("address", 5);
            resultColumnDic.Add("description", 6);
            resultColumnDic.Add("id", 7);
            resultColumnDic.Add("latitude", 8);
            resultColumnDic.Add("longitude", 9);
            resultColumnDic.Add("name", 10);
            resultColumnDic.Add("phone", 11);
            resultColumnDic.Add("promotion_info", 12);
            //resultColumnDic.Add("searchLat", 13);
            //resultColumnDic.Add("searchLng", 14);
            resultColumnDic.Add("xShard", 13);
            resultColumnDic.Add("dataText", 14);
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            return resultEW;
        }

        private void GetShopPageUrls(IListSheet listSheet)
        {
            int fileIndex = 1;

            string[] citys = this.Parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);

            for (int c = 0; c < citys.Length; c++)
            {
                string city = citys[c];
                ExcelWriter resultEW = this.CreateDetailFileWriter(city);
                Dictionary<string, string> urlDic = new Dictionary<string, string>();
                for (int i = 0; i < listSheet.RowCount; i++)
                {

                    Dictionary<string, string> row = listSheet.GetRow(i);
                    string detailPageUrl = row[SysConfig.DetailPageUrlFieldName];
                    string elemeCity = row["elemeCity"];
                    string searchLat = row["lat"];
                    string searchLng = row["lng"];
                    if (city == elemeCity)
                    {
                        string sourceDir = this.RunPage.GetDetailSourceFileDir();
                        string shopsFilePath = this.RunPage.GetFilePath(detailPageUrl, sourceDir);

                        ExcelReader er = new ExcelReader(shopsFilePath);
                        int rowCount = er.GetRowCount();
                        for (int j = 0; j < rowCount; j++)
                        {
                            Dictionary<string, string> subRow = er.GetFieldValues(j);
                            string shopId = subRow["id"];

                            if (!urlDic.ContainsKey(shopId))
                            {
                                urlDic.Add(shopId, null);

                                string lng = subRow["longitude"];
                                string lat = subRow["latitude"];

                                string dataText = "{\"timeout\":15000,\"requests\":{\"rst\":{\"method\":\"GET\",\"url\":\"/shopping/restaurant/" + shopId + "?extras[]=activities&extras[]=albums&extras[]=license&extras[]=identification&extras[]=qualification&terminal=h5&latitude=" + lat + "&longitude=" + lng + "\"},\"menu\":{\"method\":\"GET\",\"url\":\"/shopping/v2/menu?restaurant_id=" + shopId + "&terminal=h5\"},\"recommend\":{\"method\":\"GET\",\"url\":\"/shopping/v1/restaurants/" + shopId + "/quality_combo\"},\"redpack\":{\"method\":\"GET\",\"url\":\"/shopping/v1/restaurants/" + shopId + "/exclusive_hongbao/overview?code=0.29063060995754464&terminal=h5&latitude=" + lat + "&longitude=" + lng + "\"}}}";
                                string xShard = "shopid=" + shopId + ";loc=" + lng + "," + lat;

                                //string url = "https://h5.ele.me/restapi/shopping/v2/menu?restaurant_id=" + shopId + "&terminal=h5";
                                string url = "https://h5.ele.me/restapi/batch/v2?trace=shop_detail_h5&restaurant_id=" + shopId;
                                Dictionary<string, string> mapRow = new Dictionary<string, string>();
                                mapRow.Add("detailPageUrl", url);
                                mapRow.Add("detailPageName", shopId);
                                mapRow.Add("cookie", "ubt_ssid=it9p04k4zaln5h1w85u4kde2jxafztrb_2018-07-09; _utrace=6b179566de24072a5b0ccc3373d3cd38_2018-07-09; perf_ssid=oq7z6r1oglsnviwn5dd8dn6w6drjcsue_2018-07-09");
                                mapRow.Add("address", subRow["address"]);
                                mapRow.Add("description", subRow["description"]);
                                mapRow.Add("id", subRow["id"]);
                                mapRow.Add("latitude", subRow["latitude"]);
                                mapRow.Add("longitude", subRow["longitude"]);
                                mapRow.Add("name", subRow["name"]);
                                mapRow.Add("phone", subRow["phone"]);
                                mapRow.Add("promotion_info", subRow["promotion_info"]);
                                //mapRow.Add("searchLat", searchLat);
                                //mapRow.Add("searchLng", searchLng);
                                //mapRow.Add("elemeCity", elemeCity);
                                mapRow.Add("xShard", xShard);
                                mapRow.Add("dataText", dataText);
                                resultEW.AddRow(mapRow);
                            }

                        }
                    }
                }
                resultEW.SaveToDisk();
            }
        }
    }
}