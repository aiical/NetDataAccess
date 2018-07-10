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

namespace NetDataAccess.Extended.GaodeMapBoundary
{
    /// <summary>
    /// 获取各地区全名
    /// </summary>
    public class MapDistrictName : ExternalRunWebPage
    {

        #region 获取各地区全名
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            return this.CreateGetBoundaryFile(listSheet) && this.CreateGetMainPointFile(listSheet);
        }
        private bool CreateGetBoundaryFile(IListSheet listSheet)
        {
            string[] allParameters = this.Parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);

            string exportDir = allParameters[0];

            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                "detailPageUrl",
                "detailPageName", 
                "cookie",
                "grabStatus", 
                "giveUpGrab", 
                "name",
                "trimCode",
                "code",
                "fullName",
                "shortName",
                "itemIndex"});

            string resultFilePath = Path.Combine(exportDir, "高德地图行政区划边界.xlsx");

            //输出对象
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic);
            Dictionary<string, string> code2Names = new Dictionary<string, string>();
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);

                string code = row["code"];
                string name = row["name"];

                string trim0Code = code;
                while (trim0Code.EndsWith("00"))
                {
                    trim0Code = trim0Code.Substring(0, trim0Code.Length - 2);
                }
                code2Names.Add(trim0Code, name);
                string fullName = name;
                if (trim0Code.Length > 2)
                {
                    string tempCode = trim0Code;
                    while (tempCode.Length > 2)
                    {
                        tempCode = tempCode.Substring(0, tempCode.Length - 2);
                        if (code2Names.ContainsKey(tempCode))
                        {
                            string parentName = code2Names[tempCode];
                            fullName = parentName + fullName;
                        }
                    }
                }
                string shortName = name;
                if (trim0Code.Length > 2)
                {
                    string tempCode = trim0Code;
                    while (tempCode.Length > 2)
                    {
                        tempCode = tempCode.Substring(0, tempCode.Length - 2);
                        if (code2Names.ContainsKey(tempCode))
                        {
                            string parentName = code2Names[tempCode];
                            shortName = parentName + shortName;
                            break;
                        }
                    }
                }

                trim0Code = trim0Code.PadRight(6, '0');
                try
                {
                    Dictionary<string, string> f2vs = new Dictionary<string, string>();
                    f2vs.Add("detailPageUrl", row["detailPageUrl"]);
                    f2vs.Add("detailPageName", row["detailPageName"]);
                    f2vs.Add("code", code);
                    f2vs.Add("trimCode", trim0Code);
                    f2vs.Add("name", name);
                    f2vs.Add("fullName", fullName);
                    f2vs.Add("shortName", shortName);
                    f2vs.Add("itemIndex", i.ToString());
                    resultEW.AddRow(f2vs);
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }

            //保存到硬盘
            resultEW.SaveToDisk();

            return true;
        }
        private bool CreateGetMainPointFile(IListSheet listSheet)
        {
            string[] allParameters = this.Parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);

            string exportDir = allParameters[0];

            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                "detailPageUrl",
                "detailPageName", 
                "cookie",
                "grabStatus", 
                "giveUpGrab", 
                "name",
                "trimCode",
                "code",
                "fullName",
                "shortName",
                "itemIndex"});

            string resultFilePath = Path.Combine(exportDir, "高德地图行政区划点.xlsx");

            //输出对象
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic);
            Dictionary<string, string> code2Names = new Dictionary<string, string>();
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);

                string code = row["code"];
                string name = row["name"];

                string trim0Code = code;
                while (trim0Code.EndsWith("00"))
                {
                    trim0Code = trim0Code.Substring(0, trim0Code.Length - 2);
                }
                code2Names.Add(trim0Code, name);
                string fullName = name;
                if (trim0Code.Length > 2)
                {
                    string tempCode = trim0Code;
                    while (tempCode.Length > 2)
                    {
                        tempCode = tempCode.Substring(0, tempCode.Length - 2);
                        if (code2Names.ContainsKey(tempCode))
                        {
                            string parentName = code2Names[tempCode];
                            fullName = parentName + fullName;
                        }
                    }
                }
                string shortName = name;
                if (trim0Code.Length > 2)
                {
                    string tempCode = trim0Code;
                    while (tempCode.Length > 2)
                    {
                        tempCode = tempCode.Substring(0, tempCode.Length - 2);
                        if (code2Names.ContainsKey(tempCode))
                        {
                            string parentName = code2Names[tempCode];
                            shortName = parentName + shortName;
                            break;
                        }
                    }
                }

                trim0Code = trim0Code.PadRight(6, '0');
                try
                {
                    Dictionary<string, string> f2vs = new Dictionary<string, string>();
                    f2vs.Add("detailPageUrl", row["detailPageUrl"]);
                    f2vs.Add("detailPageName", row["detailPageName"]);
                    f2vs.Add("code", code);
                    f2vs.Add("trimCode", trim0Code);
                    f2vs.Add("name", name);
                    f2vs.Add("fullName", fullName);
                    f2vs.Add("shortName", shortName);
                    f2vs.Add("itemIndex", i.ToString());
                    resultEW.AddRow(f2vs);
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }

            //保存到硬盘
            resultEW.SaveToDisk();

            return true;
        }
        #endregion 
    }
}