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

namespace NetDataAccess.Extended.Jnghy.Yichuxing
{
    /// <summary>
    /// 宜出行切块
    /// </summary>
    public class MapRangeSplit : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            return this.GetRanges(this.Parameters, listSheet);
        }

        #region 获取分割后的区域
        private bool GetRanges(string parameters, IListSheet listSheet)
        {
            string[] allParameters = parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
             
            string rightStr = allParameters[0];
            string topStr = allParameters[1];
            string leftStr = allParameters[2];
            string bottomStr = allParameters[3];
            string blockWidthStr = allParameters[4];
            decimal top = decimal.Parse(topStr);
            decimal bottom = decimal.Parse(bottomStr);
            decimal left = decimal.Parse(leftStr);
            decimal right = decimal.Parse(rightStr);
            decimal blockWidth = decimal.Parse(blockWidthStr);
            string exportDir = allParameters[5]; 

            int vBlockCount = (int)Math.Ceiling((top - bottom) / blockWidth);
            int hBlockCount = (int)Math.Ceiling((right - left) / blockWidth); 
            decimal lastY = bottom; 

            //已经下载下来的列表页html保存到的目录（文件夹）
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            //输出excel表格包含的列
            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                "detailPageUrl",
                "detailPageName", 
                "cookie",
                "grabStatus", 
                "giveUpGrab", 
                "name",
                "left",
                "right",
                "bottom",
                "top",
                "itemIndex", 
                "requestString"});

            string resultFilePath = Path.Combine(exportDir, "宜出行获取POI信息.xlsx");

            //输出对象
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic);
            int itemIndex = 0;
            for (int i = 1; i <= vBlockCount; i++)
            {
                decimal nextY = i == vBlockCount ? top : (lastY + blockWidth);
                decimal lastX = left;
                for (int j = 1; j <= hBlockCount; j++)
                {
                    itemIndex++;
                    decimal nextX = j == hBlockCount ? right : (lastX + blockWidth);

                    string code =  lastX.ToString() + "," + nextX.ToString() + "," + lastY.ToString() + "," + nextY.ToString();
                    string requestString = "lat=" + nextX.ToString() + "&lng=" + nextY.ToString() + "&_token=";

                    Dictionary<string, string> f2vsLocation = new Dictionary<string, string>();
                    f2vsLocation.Add("detailPageUrl", "http://c.easygo.qq.com/api/egc/location?" + itemIndex.ToString());
                    f2vsLocation.Add("detailPageName", "location_" + code); 
                    f2vsLocation.Add("name", "location_" + code);
                    f2vsLocation.Add("left", lastX.ToString());
                    f2vsLocation.Add("right", nextX.ToString());
                    f2vsLocation.Add("bottom", lastY.ToString());
                    f2vsLocation.Add("top", nextY.ToString());
                    f2vsLocation.Add("itemIndex", itemIndex.ToString());
                    f2vsLocation.Add("requestString", requestString);
                    resultEW.AddRow(f2vsLocation);

                    Dictionary<string, string> f2vsLineData = new Dictionary<string, string>();
                    f2vsLineData.Add("detailPageUrl", "http://c.easygo.qq.com/api/egc/linedata?" + itemIndex.ToString());
                    f2vsLineData.Add("detailPageName", "lineData" + code); 
                    f2vsLineData.Add("name", "lineData" + code);
                    f2vsLineData.Add("left", lastX.ToString());
                    f2vsLineData.Add("right", nextX.ToString());
                    f2vsLineData.Add("bottom", lastY.ToString());
                    f2vsLineData.Add("top", nextY.ToString());
                    f2vsLineData.Add("itemIndex", itemIndex.ToString());
                    f2vsLineData.Add("requestString", requestString);
                    resultEW.AddRow(f2vsLineData);
                    lastX = nextX;
                }
                lastY = nextY;
            } 

            //保存到硬盘
            resultEW.SaveToDisk();

            return true;
        }
        #endregion 
    }
}