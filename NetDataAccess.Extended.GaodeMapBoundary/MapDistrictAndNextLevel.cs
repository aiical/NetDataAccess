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

namespace NetDataAccess.Extended.GaodeMapBoundary
{
    /// <summary>
    /// 获取各地区及下级读取构成的js boundary信息
    /// </summary>
    public class MapDistrictAndNextLevel : ExternalRunWebPage
    { 
        private Dictionary<string, Dictionary<string, string>> _DistrictDic = null;
        private Dictionary<string, Dictionary<string, string>> DistrictDic
        {
            get
            {
                return this._DistrictDic;
            }
        }
        
        public override bool BeforeAllGrab()
        {
            string[] parts = this.Parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            string districtListFilePath = parts[0];

            CsvReader er = new CsvReader(districtListFilePath);
            int rowCount = er.GetRowCount();
            Dictionary<string, Dictionary<string, string>> districtDic = new Dictionary<string, Dictionary<string, string>>();
            for (int i = 0; i < rowCount; i++)
            {
                Dictionary<string, string> districtInfo = er.GetFieldValues(i);
                string code = districtInfo["code"];
                districtDic.Add(code, districtInfo);
            }
            this._DistrictDic = districtDic;
            return base.BeforeAllGrab();
        }

        public override void GetDataByOtherAcessType(Dictionary<string, string> listRow)
        {
            string code = listRow["code"];
            string detailPageUrl = listRow["detailPageUrl"]; 
            string trimCode = code;
            while (trimCode.EndsWith("00"))
            {
                trimCode = trimCode.Substring(0, trimCode.Length - 2);
            }
            List<string> childCodes = this.GetNextLevelDistrictCodes(trimCode);
            if (childCodes.Count == 0 && trimCode.Length == 2)
            {
                //直辖市
                childCodes = this.GetNextLevelDistrictCodes(trimCode + "01");
            }
            decimal maxX = 0;
            decimal maxY = 0;
            decimal minX = decimal.MaxValue;
            decimal minY = decimal.MaxValue;

            Dictionary<string, string> districtInfo = this.DistrictDic[code];
            JObject mainJsonObj = new JObject();
            string stringFormat = "#0.0";
            int pointZipSize = 30;
            string pointZipStringFormat = "#0.00";
            mainJsonObj.Add("code", code);
            mainJsonObj.Add("name", districtInfo["name"]);
            JArray boundaryArray = this.GetDistrictBoundaryArray(districtInfo["boundaryPoints"], stringFormat, pointZipSize, pointZipStringFormat, out maxX, out maxY, out minX, out minY);
            if (maxY - minY < (decimal)0.5 || maxX - minX < (decimal)0.5)
            {
                pointZipStringFormat = "#0.00000";
                stringFormat = "#0.00000";
                boundaryArray = this.GetDistrictBoundaryArray(districtInfo["boundaryPoints"], stringFormat, pointZipSize, pointZipStringFormat, out maxX, out maxY, out minX, out minY);
            }
            else if (maxY - minY < (decimal)2 || maxX - minX < (decimal)2)
            {
                pointZipStringFormat = "#0.000";
                stringFormat = "#0.000";
                boundaryArray = this.GetDistrictBoundaryArray(districtInfo["boundaryPoints"], stringFormat, pointZipSize, pointZipStringFormat, out maxX, out maxY, out minX, out minY);
            }
            else if (maxY - minY < 20 || maxX - minX < 20)
            {
                stringFormat = "#0.00";
                boundaryArray = this.GetDistrictBoundaryArray(districtInfo["boundaryPoints"], stringFormat, pointZipSize, pointZipStringFormat, out maxX, out maxY, out minX, out minY);
            } 
            mainJsonObj.Add("boundaryArray", boundaryArray);
            mainJsonObj.Add("maxX", maxX);
            mainJsonObj.Add("maxY", maxY);
            mainJsonObj.Add("minX", minX);
            mainJsonObj.Add("minY", minY);

            JArray nextLevelArray = new JArray();
            mainJsonObj.Add("nextLevelArray", nextLevelArray);
            StringBuilder ss = new StringBuilder("boundaryList[\"" + code + "_L1\"] = ");

            for (int i = 0; i < childCodes.Count; i++)
            {
                nextLevelArray.Add(this.GetDistrictJson(childCodes[i], stringFormat, pointZipSize, pointZipStringFormat));
            }

            ss.AppendLine(mainJsonObj.ToString()); 

            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            string localFilePath = this.RunPage.GetFilePath(detailPageUrl+"_L1.js", pageSourceDir);

            FileHelper.SaveTextToFile(ss.ToString(), localFilePath);

        }

        private JObject GetDistrictJson(string code, string stringFormat, int pointZipSize, string pointZipStringFormat)
        {
            decimal maxX = 0;
            decimal maxY = 0;
            decimal minX = decimal.MaxValue;
            decimal minY = decimal.MaxValue;
            Dictionary<string, string> districtInfo = this.DistrictDic[code];
            JObject jsonObj = new JObject();
            jsonObj.Add("code", code);
            jsonObj.Add("name", districtInfo["name"]);
            jsonObj.Add("boundaryArray", this.GetDistrictBoundaryArray(districtInfo["boundaryPoints"], stringFormat, pointZipSize, pointZipStringFormat, out maxX, out maxY, out minX, out minY));
            return jsonObj;
        }

        private JArray GetDistrictBoundaryArray(string boundaryStr, string stringFormat,int pointZipSize, string pointZipStringFormat,  out decimal maxX, out decimal maxY, out decimal minX, out decimal minY)
        {
            maxX = 0;
            maxY = 0;
            minX = decimal.MaxValue;
            minY = decimal.MaxValue;
            JArray boundaryArray = new JArray();
            string[] parts = boundaryStr.Split(new string[] { "|" }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string part in parts)
            {
                if (part.Trim().Length != 0)
                {
                    List<decimal[]> pointList = new List<decimal[]>();
                    string[] values = part.Trim().Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                    for (int i = 0; i < values.Length; i = i + 2)
                    {
                        pointList.Add(new decimal[] { decimal.Parse(values[i].Trim()), decimal.Parse(values[i + 1].Trim()) });
                    }
                    StringBuilder s = new StringBuilder();
                    string tempStr = "";
                    int pointCount = 0;
                    string tempStringFormat = pointList.Count < pointZipSize ? pointZipStringFormat : stringFormat;
                    foreach (decimal[] point in pointList)
                    {
                        string str = point[0].ToString(tempStringFormat) + "," + point[1].ToString(tempStringFormat) + " ";
                        if (str != tempStr)
                        {
                            if (point[0] > maxX)
                            {
                                maxX = point[0];
                            }

                            if (point[1] > maxY)
                            {
                                maxY = point[1];
                            }

                            if (point[0] < minX)
                            {
                                minX = point[0];
                            }

                            if (point[1] < minY)
                            {
                                minY = point[1];
                            }

                            tempStr = str;
                            s.Append(str);
                            pointCount++;
                        }
                    }
                    if (pointCount > 3)
                    {
                        boundaryArray.Add(s.ToString());
                    }
                }
            }
            return boundaryArray;
        }

        private List<string> GetNextLevelDistrictCodes(string trimCode)
        {
            List<string> codeList = new List<string>();
            foreach (string key in this.DistrictDic.Keys)
            {
                if (trimCode.PadRight(6, '0') != key)
                {
                    if (trimCode == "中国")
                    {
                        if (key.EndsWith("0000"))
                        {
                            codeList.Add(key);
                        }
                    }
                    else if (trimCode.Length == 2)
                    {
                        if (key.StartsWith(trimCode) && key.EndsWith("00"))
                        {
                            codeList.Add(key);
                        }
                    }
                    else if (trimCode.Length == 4)
                    {
                        if (key.StartsWith(trimCode))
                        {
                            codeList.Add(key);
                        }
                    }
                }
            }
            return codeList;
        }

        #region 获取各地区及下级读取构成的js boundary信息
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            return true;
        }
        #endregion 
    }
}