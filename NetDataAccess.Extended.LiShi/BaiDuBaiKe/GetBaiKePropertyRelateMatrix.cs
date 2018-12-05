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

namespace NetDataAccess.Extended.LiShi.BaiDuBaiKe
{
    public class GetBaiKePropertyRelateMatrix : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            this.GetPropertiesMatrix(listSheet);
            return true;
        }

        private void GetPropertiesMatrix(IListSheet listSheet)
        {
            string[] parameters = this.Parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            string sourceFilePath = parameters[0];
            string destFilePath = parameters[1];

            ExcelReader er = new ExcelReader(sourceFilePath);
            int sourceRowCount = er.GetRowCount();

            Dictionary<string, int> allPropertyCountDic = new Dictionary<string, int>();

            List<string> allPropertyList = new List<string>();
            for (int i = 0; i < sourceRowCount; i++)
            {
                Dictionary<string,string> sourceRow = er.GetFieldValues(i);
                string[] itemProperties = sourceRow["properties"].Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string itemProperty in itemProperties)
                { 
                    if (allPropertyCountDic.ContainsKey(itemProperty))
                    {
                        allPropertyCountDic[itemProperty] = allPropertyCountDic[itemProperty] + 1;
                    }
                    else
                    {
                        allPropertyList.Add(itemProperty);
                        allPropertyCountDic.Add(itemProperty, 1);
                    }
                }
            }

            //如果出现少于等于2次，那么忽略此属性
            int ignoreNum = 6;
            List<string> propertyList = new List<string>();
            Dictionary<string, bool> propertyListDic = new Dictionary<string, bool>();
            foreach (string itemProperty in allPropertyList)
            {
                if (allPropertyCountDic[itemProperty] > ignoreNum)
                {
                    propertyList.Add(itemProperty);
                    propertyListDic.Add(itemProperty, true);
                }
            }

            int maxTime = 1;

            Dictionary<string, Dictionary<string, int>> pToPDic = new Dictionary<string, Dictionary<string, int>>();
            for (int i = 0; i < sourceRowCount; i++)
            {
                Dictionary<string, string> sourceRow = er.GetFieldValues(i);
                string[] itemProperties = sourceRow["properties"].Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string fromItemProperty in itemProperties)
                {
                    if (propertyListDic.ContainsKey(fromItemProperty))
                    {
                        if (!pToPDic.ContainsKey(fromItemProperty))
                        {
                            pToPDic.Add(fromItemProperty, new Dictionary<string, int>());
                        }
                        Dictionary<string, int> propertyDic = pToPDic[fromItemProperty];

                        if (!propertyDic.ContainsKey(fromItemProperty))
                        {
                            propertyDic.Add(fromItemProperty, 1);
                        }
                        else
                        {
                            propertyDic[fromItemProperty] = propertyDic[fromItemProperty] + 1;
                        }

                        foreach (string toItemProperty in itemProperties)
                        {
                            if (propertyListDic.ContainsKey(toItemProperty))
                            {
                                if (fromItemProperty != toItemProperty)
                                {
                                    if (!propertyDic.ContainsKey(toItemProperty))
                                    {
                                        propertyDic.Add(toItemProperty, 1);
                                    }
                                    else
                                    {
                                        int tmpValue = propertyDic[toItemProperty] + 1;
                                        propertyDic[toItemProperty] = tmpValue;
                                        if (tmpValue > maxTime)
                                        {
                                            maxTime = tmpValue;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }


            
            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("pToP", 0);
            for (int i = 0; i < propertyList.Count; i++)
            {
                resultColumnDic.Add(propertyList[i], i + 1);
            }
            
            CsvWriter propertyMatrixCW = new CsvWriter(destFilePath, resultColumnDic);

            foreach (string fromProperty in propertyList)
            { 

                Dictionary<string, string> resultRow = new Dictionary<string,string>();
                resultRow.Add("pToP", fromProperty);
                Dictionary<string, int> propertyDic = pToPDic.ContainsKey(fromProperty) ? pToPDic[fromProperty] : null;
                foreach (string toProperty in propertyList)
                {
                    double value = fromProperty == toProperty ? 0 : (propertyDic == null || !propertyDic.ContainsKey(toProperty) || propertyDic[toProperty] == 0 ? 2 * (double)maxTime : ((double)maxTime / (double)propertyDic[toProperty]));
                    resultRow.Add(toProperty, value.ToString());
                }
                propertyMatrixCW.AddRow(resultRow);
            }

            propertyMatrixCW.SaveToDisk();

            string allPropertyNameFilePath = destFilePath + "_AllPropertyName.xlsx";
            Dictionary<string, int> allPropertyNameColumnDic = new Dictionary<string, int>();
            allPropertyNameColumnDic.Add("name", 0);
            allPropertyNameColumnDic.Add("count", 1);
            Dictionary<string, string> allPropertyNameColumnFormats = new Dictionary<string, string>();
            allPropertyNameColumnFormats.Add("count", "#0");
            ExcelWriter allPropertyNameEW = new ExcelWriter(allPropertyNameFilePath, "List", allPropertyNameColumnDic, allPropertyNameColumnFormats);
            for (int i = 0; i < allPropertyList.Count; i++)
            {
                string fromProperty = allPropertyList[i];
                Dictionary<string, object> resultRow = new Dictionary<string, object>();
                resultRow.Add("name", fromProperty);
                resultRow.Add("count", allPropertyCountDic[fromProperty]);
                allPropertyNameEW.AddRow(resultRow);

            }
            allPropertyNameEW.SaveToDisk();

            string propertyNameFilePath = destFilePath + "_PropertyName.xlsx";
            Dictionary<string, int> propertyNameColumnDic = new Dictionary<string, int>();
            propertyNameColumnDic.Add("name", 0);
            ExcelWriter propertyNameEW = new ExcelWriter(propertyNameFilePath, "List", propertyNameColumnDic);
            for (int i = 0; i < propertyList.Count; i++)
            {
                string fromProperty = propertyList[i];
                Dictionary<string, string> resultRow = new Dictionary<string, string>();
                resultRow.Add("name", fromProperty);
                propertyNameEW.AddRow(resultRow);

            }
            propertyNameEW.SaveToDisk();


            string propertyArrayFilePath = destFilePath + "_Array.txt";
            StringBuilder propertyArrayStringBuilder = new StringBuilder();
            propertyArrayStringBuilder.Append("arr = [");
            for (int i = 0; i < propertyList.Count; i++)
            {
                string fromProperty = propertyList[i];
                propertyArrayStringBuilder.Append((i == 0 ? "" : ", \r\n") + "[");
                Dictionary<string, string> resultRow = new Dictionary<string, string>();
                resultRow.Add("pToP", fromProperty);
                Dictionary<string, int> propertyDic = pToPDic.ContainsKey(fromProperty) ? pToPDic[fromProperty] : null;
                for (int j = 0; j < propertyListDic.Count; j++)
                {
                    string toProperty = propertyList[j];
                    double value = fromProperty == toProperty ? 0 : (propertyDic == null || !propertyDic.ContainsKey(toProperty) || propertyDic[toProperty] == 0 ? 2 * (double)maxTime : ((double)maxTime / (double)propertyDic[toProperty]));
                    resultRow.Add(toProperty, value.ToString());
                    propertyArrayStringBuilder.Append((j == 0 ? "" : ", ") + value.ToString());
                }
                propertyMatrixCW.AddRow(resultRow);
                propertyArrayStringBuilder.Append("]");
            }
            propertyArrayStringBuilder.Append("]");
            FileHelper.SaveTextToFile(propertyArrayStringBuilder.ToString(), propertyArrayFilePath);
        } 
    }
}