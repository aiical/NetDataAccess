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

namespace NetDataAccess.Extended.Baidu
{
    /// <summary>
    /// 车维修获取列表页
    /// </summary>
    public class MapRangeSplit : ExternalRunWebPage
    { 

        #region 获取分割后的区域
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            string[] allParameters = this.Parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);

            string keyWords = allParameters[0];
            string rightStr = allParameters[1];
            string topStr = allParameters[2];
            string leftStr = allParameters[3];
            string bottomStr = allParameters[4];
            string blockWidthStr = allParameters[5];
            decimal top = decimal.Parse(topStr);
            decimal bottom = decimal.Parse(bottomStr);
            decimal left = decimal.Parse(leftStr);
            decimal right = decimal.Parse(rightStr);
            decimal blockWidth = decimal.Parse(blockWidthStr);
            string exportDir = allParameters[6];

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
                "keyWords"});

            string resultFilePath = Path.Combine(exportDir, "百度地图搜索.xlsx");

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

                    string code = lastX.ToString() + "," + nextX.ToString() + "," + lastY.ToString() + "," + nextY.ToString() + "," + keyWords;

                    Dictionary<string, string> f2vs = new Dictionary<string, string>();
                    f2vs.Add("detailPageUrl", code);
                    f2vs.Add("detailPageName", code);
                    f2vs.Add("name", code);
                    f2vs.Add("left", lastX.ToString());
                    f2vs.Add("right", nextX.ToString());
                    f2vs.Add("bottom", lastY.ToString());
                    f2vs.Add("top", nextY.ToString());
                    f2vs.Add("itemIndex", itemIndex.ToString());
                    f2vs.Add("keyWords", keyWords);
                    resultEW.AddRow(f2vs);
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