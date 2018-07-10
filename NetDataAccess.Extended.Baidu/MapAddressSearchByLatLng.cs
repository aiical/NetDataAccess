using NetDataAccess.Base.Common;
using NetDataAccess.Base.DB;
using NetDataAccess.Base.DLL;
using NetDataAccess.Base.Writer;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace NetDataAccess.Extended.Baidu
{
    public class MapAddressSearchByLatLng : CustomProgramBase
    {
        #region 入口函数
        public bool Run(string parameters, IListSheet listSheet)
        {
            return GetAddress(parameters, listSheet);
        }
        #endregion

        #region GetAddress
        private bool GetAddress(string parameters, IListSheet listSheet)
        {
            //已经下载下来的首页html保存到的目录（文件夹）
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("uid", 0);
            resultColumnDic.Add("title", 1);
            resultColumnDic.Add("address", 2);
            resultColumnDic.Add("province", 3);
            resultColumnDic.Add("city", 4);
            resultColumnDic.Add("district", 5);
            resultColumnDic.Add("street", 6);
            resultColumnDic.Add("streetNumber", 7);
            resultColumnDic.Add("adcode", 8); 
            resultColumnDic.Add("phoneNumber", 9);
            resultColumnDic.Add("postcode", 10);
            resultColumnDic.Add("lat", 11);
            resultColumnDic.Add("lng", 12);
            resultColumnDic.Add("detailUrl", 13);
            resultColumnDic.Add("url", 14);
            string exportDir = this.RunPage.GetExportDir();

            string resultFilePath = Path.Combine(exportDir, "百度地图_包括行政区街道门牌号.xlsx");

            //输出对象
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic);

            //从下载到的首页html中，获取列表页个数，并形成所有列表页url
            GetAddress(listSheet, pageSourceDir, resultEW);

            //保存到硬盘
            resultEW.SaveToDisk();

            return true;
        }
        #endregion

        #region GetAddress
        /// <summary>
        /// GetAddress
        /// </summary>
        /// <param name="listSheet"></param>
        /// <param name="pageSourceDir"></param>
        /// <param name="resultEW"></param>
        private void GetAddress(IListSheet listSheet, string pageSourceDir, ExcelWriter resultEW)
        {
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                //listSheet中只有一条记录
                string pageUrl = listSheet.PageUrlList[i];
                Dictionary<string, string> row = listSheet.GetRow(i);

                string localFilePath = this.RunPage.GetFilePath(pageUrl, pageSourceDir); 
                string fileText = FileHelper.GetTextFromFile(localFilePath);
                int jsonBeginIndex = fileText.IndexOf("{");
                int jsonEndIndex = fileText.LastIndexOf("}");
                string jsonStr = fileText.Substring(jsonBeginIndex, jsonEndIndex - jsonBeginIndex + 1);
                JObject rootJo = JObject.Parse(jsonStr);
                JObject addressJo = rootJo.SelectToken("result").SelectToken("addressComponent") as JObject;
                string adcode = addressJo.GetValue("adcode") == null ? "" : addressJo.GetValue("adcode").ToString();
                string district = addressJo.GetValue("district") == null ? "" : addressJo.GetValue("district").ToString();
                string street = addressJo.GetValue("street") == null ? "" : addressJo.GetValue("street").ToString();
                string streetNumber = addressJo.GetValue("street_number") == null ? "" : addressJo.GetValue("street_number").ToString();

                Dictionary<string, string> p2vs = new Dictionary<string, string>();
                p2vs.Add("address", row["address"]);
                p2vs.Add("city", row["city"]);
                p2vs.Add("detailUrl", row["detailUrl"]);
                p2vs.Add("isAccurate", row["isAccurate"]);
                p2vs.Add("phoneNumber", row["phoneNumber"]);
                p2vs.Add("postcode", row["postcode"]);
                p2vs.Add("province", row["province"]);
                p2vs.Add("title", row["title"]);
                p2vs.Add("uid", row["uid"]);
                p2vs.Add("url", row["url"]);
                p2vs.Add("lat", row["lat"]);
                p2vs.Add("lng", row["lng"]);
                p2vs.Add("adcode", adcode);
                p2vs.Add("disctrict", district);
                p2vs.Add("street", street);
                p2vs.Add("streetNumber", streetNumber);

                resultEW.AddRow(p2vs);
            }
        }
        #endregion
    }
}
