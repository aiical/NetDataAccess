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

namespace NetDataAccess.Extended.Gaode
{
    public class MapPointSplit : ExternalRunWebPage
    {

        #region 获取等距离选择的点坐标
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            string[] allParameters = this.Parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);

            string minXStr = allParameters[0];
            string maxXStr = allParameters[1];
            string minYStr = allParameters[2];
            string maxYStr = allParameters[3];
            string xPointCountStr = allParameters[4];
            string yPointCountStr = allParameters[5];
            decimal minX = decimal.Parse(minXStr);
            decimal minY = decimal.Parse(minYStr);
            decimal maxX = decimal.Parse(maxXStr);
            decimal maxY = decimal.Parse(maxYStr);
            decimal xPointCount = decimal.Parse(xPointCountStr);
            decimal yPointCount = decimal.Parse(yPointCountStr);
            string exportDir = allParameters[6];
            int fileIndex = 1;

            ExcelWriter ew = null;
            decimal xPerWith = (maxX - minX) / xPointCount;
            decimal yPerWith = (maxY - minY) / yPointCount;
            for (int i = 0; i <= xPointCount; i++)
            {
                decimal x = minX + xPerWith * i;
                for (int j = 0; j <= yPointCount; j++)
                {
                    decimal y = minY + yPerWith * j;

                    if (ew == null || ew.RowCount >= 500000)
                    {
                        if (ew != null)
                        {
                            ew.SaveToDisk();
                        }
                        ew = this.GetPointWirter(exportDir, fileIndex);
                        fileIndex++;
                    }
                    string xStr = x.ToString("#0.000000");
                    string yStr = y.ToString("#0.000000");
                    Dictionary<string, string> f2vs = new Dictionary<string, string>();
                    f2vs.Add("detailPageUrl", xStr + "_" + yStr);
                    f2vs.Add("detailPageName", xStr + "_" + yStr);
                    f2vs.Add("lat", xStr);
                    f2vs.Add("lng", yStr);
                    ew.AddRow(f2vs);
                }
            }

            //保存到硬盘
            ew.SaveToDisk();

            return true;
        }
        #endregion

        private ExcelWriter GetPointWirter(string exportDir, int fileIndex)
        {
            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                "detailPageUrl",
                "detailPageName", 
                "cookie",
                "grabStatus", 
                "giveUpGrab",  
                "lat",
                "lng"});

            string resultFilePath = Path.Combine(exportDir, "高德地图获取点信息_" + fileIndex.ToString() + ".xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic);
            return resultEW;
        }
    }
}