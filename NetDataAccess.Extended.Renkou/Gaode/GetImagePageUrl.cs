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

namespace NetDataAccess.Extended.Renkou.Gaode
{
    public class GetImagePageUrl : CustomProgramBase
    {
        public bool Run(string parameters, IListSheet listSheet)
        {
            return GetFirstListPageFromPage(parameters, listSheet);
        } 

        private int ToGaodeX(double baiduX, int z)
        {
            switch (z)
            {
                case 18:
                    return (int)((baiduX + 20038389.7977) / 152.88274);
                default:
                    throw new Exception("无法处理坐标, z=" + z.ToString());
            }
        }

        private int ToGaodeY(double baiduY, int z)
        {
            switch (z)
            {
                case 18:
                    return (int)((baiduY - 19952702.3442) / -152.2894);
                default:
                    throw new Exception("无法处理坐标, z=" + z.ToString());
            }
        }

        private bool GetFirstListPageFromPage(string parameters, IListSheet listSheet)
        {
            string[] ps = parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            int z = int.Parse(ps[4]);
            int left = ToGaodeX(double.Parse(ps[0]), z);
            int top = ToGaodeY(double.Parse(ps[1]), z);
            int right = ToGaodeX(double.Parse(ps[2]), z);
            int bottom = ToGaodeY(double.Parse(ps[3]), z);
            DateTime fromTime = DateTime.ParseExact(ps[5], "yyyyMMddHH", System.Globalization.CultureInfo.CurrentCulture);
            DateTime ToTime = DateTime.ParseExact(ps[6], "yyyyMMddHH", System.Globalization.CultureInfo.CurrentCulture);
            string exportDir = ps[7]; 

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("x", 5);
            resultColumnDic.Add("y", 6);
            resultColumnDic.Add("z", 7);
            resultColumnDic.Add("time", 8);
            string resultFilePath = Path.Combine(exportDir, "高德热力图爬取明细.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);

            while (fromTime <= ToTime)
            {
                string timeStr = fromTime.ToString("yyyyMMddHH");
                for (int x = left; x <= right; x++)
                {
                    for (int y = top; y <= bottom; y++)
                    {
                        string detailPageUrl = "http://heatmap.amap.com/api/showmap/pvheatmap?x=" + x.ToString() +"&y=" + y.ToString() + "&z=" + z.ToString() + "&showmap=equal&htime=" + timeStr;
                        string detailPageName = x.ToString() + "," + y.ToString() + "," + z.ToString() + "," + timeStr;
                        Dictionary<string, string> f2vs = new Dictionary<string, string>();
                        f2vs.Add("detailPageUrl", detailPageUrl);
                        f2vs.Add("detailPageName", detailPageName);
                        f2vs.Add("x", x.ToString());
                        f2vs.Add("y", y.ToString());
                        f2vs.Add("z", z.ToString());
                        f2vs.Add("time", timeStr);
                        resultEW.AddRow(f2vs);
                    }
                }
                fromTime = fromTime.AddHours(1);
            } 

            resultEW.SaveToDisk();

            return true;
        } 
    }
}