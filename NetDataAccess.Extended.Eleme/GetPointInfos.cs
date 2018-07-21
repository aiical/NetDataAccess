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
    public class GetPointInfos : ExternalRunWebPage
    {

        #region 获取等距离选择的点坐标
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            this.GetAllPointInfos(listSheet);
            return true;
        }
        #endregion 

        public override void CheckRequestCompleteFile(string webPageText, Dictionary<string, string> listRow)
        {
            string detailPageUrl = listRow[SysConfig.DetailPageUrlFieldName];
            string address = listRow["formattedAddress"];
            string city = listRow["city"];
            double lat = double.Parse( listRow["lng"]);
            double lng = double.Parse(listRow["lat"]);
            JArray rootJA = JArray.Parse(webPageText);
            return;

            for (int i = 0; i < rootJA.Count; i++)
            {
                JObject jo = rootJA[i] as JObject;
                string elemeAddress = jo.GetValue("address").ToString();
                string elemeCity = jo.GetValue("city").ToString();
                //double elemeLat = double.Parse(jo.GetValue("latitude").ToString());
                //double elemeLng = double.Parse(jo.GetValue("longitude").ToString());

                //名字一样
                //if (elemeCity == city && elemeAddress.Substring(elemeAddress.Length - 3) == address.Substring(address.Length - 3))
                if(elemeCity == city)
                {
                    return;
                }
            }
            /*
            for (int i = 0; i < rootJA.Count; i++)
            {
                JObject jo = rootJA[i] as JObject;
                string elemeAddress = jo.GetValue("address").ToString();
                string elemeCity = jo.GetValue("city").ToString();
                double elemeLat = double.Parse(jo.GetValue("latitude").ToString());
                double elemeLng = double.Parse(jo.GetValue("longitude").ToString());

                //距离0.05以内
                if (elemeCity == city && Math.Abs(elemeLat - lat) < 0.05 && Math.Abs(elemeLng - lng) < 0.05)
                {
                    return;
                }
            }*/
            throw new Exception("未找到匹配的地址, " + detailPageUrl);
        }

        private void GetAllPointInfos(IListSheet listSheet)
        {
            ExcelWriter ew = null;
            string sourceDir = this.RunPage.GetDetailSourceFileDir();
            int fileIndex = 1;
            int pointCount = 0;
            int rowCount = listSheet.GetListDBRowCount();
            for (int i = 0; i < rowCount; i++)
            {
                if (ew == null)
                {
                    ew = this.GetWirter(fileIndex);
                    pointCount = 0;
                    fileIndex++;
                }
                else if (pointCount >= 500000)
                {
                    ew.SaveToDisk();
                    ew = this.GetWirter(fileIndex);
                    pointCount = 0;
                    fileIndex++;
                }

                Dictionary<string, string> listRow = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(listRow[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string detailPageUrl = listRow[SysConfig.DetailPageUrlFieldName];
                    string filePath = this.RunPage.GetFilePath(detailPageUrl, sourceDir);
                    string province = listRow["province"];
                    string city = listRow["city"];
                    string address = listRow["formattedAddress"];
                    double lat = double.Parse(listRow["lng"]);
                    double lng = double.Parse(listRow["lat"]);

                    string jsonText = FileHelper.GetTextFromFile(filePath);
                    JArray rootJA = JArray.Parse(jsonText);
                    JObject pointJo = null;
                    for (int j = 0; j < rootJA.Count; j++)
                    {
                        JObject jo = rootJA[j] as JObject;
                        string elemeAddress = jo.GetValue("address").ToString();
                        string elemeCity = jo.GetValue("city").ToString();
                        double elemeLat = double.Parse(jo.GetValue("latitude").ToString());
                        double elemeLng = double.Parse(jo.GetValue("longitude").ToString());
                        if (elemeAddress.Substring(elemeAddress.Length - 3) == address.Substring(address.Length - 3))
                        {
                            pointJo = jo;
                            break;
                        }
                    }
                    if (pointJo == null)
                    {
                        for (int j = 0; j < rootJA.Count; j++)
                        {
                            JObject jo = rootJA[j] as JObject;
                            string elemeAddress = jo.GetValue("address").ToString();
                            string elemeCity = jo.GetValue("city").ToString();
                            double elemeLat = double.Parse(jo.GetValue("latitude").ToString());
                            double elemeLng = double.Parse(jo.GetValue("longitude").ToString());
                            if (Math.Abs(elemeLat - lat) < 0.05 && Math.Abs(elemeLng - lng) < 0.05)
                            {
                                pointJo = jo;
                                break;
                            }
                        }
                    }

                    if (pointJo != null)
                    {
                        string elemeAddress = pointJo.GetValue("address").ToString();
                        string elemeCity = pointJo.GetValue("city").ToString();
                        double elemeLat = double.Parse(pointJo.GetValue("latitude").ToString());
                        double elemeLng = double.Parse(pointJo.GetValue("longitude").ToString());
                        string geohash = pointJo.GetValue("geohash").ToString();
                        string short_address = pointJo.GetValue("short_address").ToString();

                        Dictionary<string, string> elemePoint = new Dictionary<string, string>();
                        elemePoint.Add("detailPageUrl", lat + "_" + lng);
                        elemePoint.Add("detailPageName", lat + "_" + lng);
                        elemePoint.Add("province", province);
                        elemePoint.Add("city", city);
                        elemePoint.Add("elemeCity", elemeCity);
                        elemePoint.Add("address", address);
                        elemePoint.Add("lat", elemeLat.ToString());
                        elemePoint.Add("lng", elemeLng.ToString());
                        elemePoint.Add("urlFormat", "https://h5.ele.me/restapi/shopping/v3/restaurants?latitude=##lat##&longitude=##lng##&keyword=&offset=##offset##&limit=8&extras[]=activities&extras[]=tags&terminal=h5&rank_id=7ab4ea35dfbd4806b6ed02cbe08ce02a&order_by=5&brand_ids[]=&restaurant_category_ids[]=209&restaurant_category_ids[]=212&restaurant_category_ids[]=213&restaurant_category_ids[]=214&restaurant_category_ids[]=215&restaurant_category_ids[]=216&restaurant_category_ids[]=217&restaurant_category_ids[]=219&restaurant_category_ids[]=265&restaurant_category_ids[]=266&restaurant_category_ids[]=267&restaurant_category_ids[]=268&restaurant_category_ids[]=269&restaurant_category_ids[]=221&restaurant_category_ids[]=222&restaurant_category_ids[]=223&restaurant_category_ids[]=224&restaurant_category_ids[]=225&restaurant_category_ids[]=226&restaurant_category_ids[]=227&restaurant_category_ids[]=228&restaurant_category_ids[]=231&restaurant_category_ids[]=232&restaurant_category_ids[]=263&restaurant_category_ids[]=218&restaurant_category_ids[]=234&restaurant_category_ids[]=235&restaurant_category_ids[]=236&restaurant_category_ids[]=237&restaurant_category_ids[]=238&restaurant_category_ids[]=211&restaurant_category_ids[]=229&restaurant_category_ids[]=230&restaurant_category_ids[]=264");
                        ew.AddRow(elemePoint);

                    }
                }
            }
            ew.SaveToDisk();
        } 

        private ExcelWriter GetWirter(int fileIndex)
        {
            string exportDir = this.RunPage.GetExportDir();
            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                "detailPageUrl",
                "detailPageName", 
                "cookie",
                "grabStatus", 
                "giveUpGrab",
                "province",
                "city", 
                "elemeCity", 
                "address", 
                "lat",
                "lng",
                "urlFormat"});

            string resultFilePath = Path.Combine(exportDir, "饿了么_店铺信息_" + fileIndex.ToString() + ".xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic);
            return resultEW;
        }
    }
}