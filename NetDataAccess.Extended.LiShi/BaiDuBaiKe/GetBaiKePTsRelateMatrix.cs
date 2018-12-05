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
    public class GetBaiKePTsRelateMatrix : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            this.GetPTsMatrix(listSheet);
            return true;
        }

        private void GetPTsMatrix(IListSheet listSheet)
        {
            string[] parameters = this.Parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            string sourceFilePath = parameters[0];
            string destFilePath = parameters[1];

            ExcelReader er = new ExcelReader(sourceFilePath);
            int sourceRowCount = er.GetRowCount();

            Dictionary<string, int> allPTCountDic = new Dictionary<string, int>();

            List<string> allPTList = new List<string>();
            for (int i = 0; i < sourceRowCount; i++)
            {
                Dictionary<string,string> sourceRow = er.GetFieldValues(i);
                string[] itemPTs = sourceRow["pts"].Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string itemPT in itemPTs)
                { 
                    if (allPTCountDic.ContainsKey(itemPT))
                    {
                        allPTCountDic[itemPT] = allPTCountDic[itemPT] + 1;
                    }
                    else
                    {
                        allPTList.Add(itemPT);
                        allPTCountDic.Add(itemPT, 1);
                    }
                }
            }

            //如果出现少于等于2次，那么忽略此属性
            int ignoreNum = 10;
            List<string> ptList = new List<string>();
            Dictionary<string, bool> ptListDic = new Dictionary<string, bool>();
            foreach (string itemPT in allPTList)
            {
                if (allPTCountDic[itemPT] > ignoreNum)
                {
                    ptList.Add(itemPT);
                    ptListDic.Add(itemPT, true);
                }
            }

            int maxTime = 1;

            Dictionary<string, Dictionary<string, int>> ptToPTDic = new Dictionary<string, Dictionary<string, int>>();
            for (int i = 0; i < sourceRowCount; i++)
            {
                Dictionary<string, string> sourceRow = er.GetFieldValues(i);
                string[] itemPTs = sourceRow["pts"].Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string fromItemPT in itemPTs)
                {
                    if (ptListDic.ContainsKey(fromItemPT))
                    {
                        if (!ptToPTDic.ContainsKey(fromItemPT))
                        {
                            ptToPTDic.Add(fromItemPT, new Dictionary<string, int>());
                        }
                        Dictionary<string, int> ptDic = ptToPTDic[fromItemPT];

                        if (!ptDic.ContainsKey(fromItemPT))
                        {
                            ptDic.Add(fromItemPT, 1);
                        }
                        else
                        {
                            ptDic[fromItemPT] = ptDic[fromItemPT] + 1;
                        }

                        foreach (string toItemPT in itemPTs)
                        {
                            if (ptListDic.ContainsKey(toItemPT))
                            {
                                if (fromItemPT != toItemPT)
                                {
                                    if (!ptDic.ContainsKey(toItemPT))
                                    {
                                        ptDic.Add(toItemPT, 1);
                                    }
                                    else
                                    {
                                        int tmpValue = ptDic[toItemPT] + 1;
                                        ptDic[toItemPT] = tmpValue;
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
            resultColumnDic.Add("ptToPT", 0);
            for (int i = 0; i < ptList.Count; i++)
            {
                resultColumnDic.Add(ptList[i], i + 1);
            }
            
            CsvWriter ptMatrixCW = new CsvWriter(destFilePath, resultColumnDic);

            foreach (string fromPT in ptList)
            { 

                Dictionary<string, string> resultRow = new Dictionary<string,string>();
                resultRow.Add("ptToPT", fromPT);
                Dictionary<string, int> propertyDic = ptToPTDic.ContainsKey(fromPT) ? ptToPTDic[fromPT] : null;
                foreach (string toPT in ptList)
                {
                    double value = fromPT == toPT ? 0 : (propertyDic == null || !propertyDic.ContainsKey(toPT) || propertyDic[toPT] == 0 ? 2 * (double)maxTime : ((double)maxTime / (double)propertyDic[toPT]));
                    resultRow.Add(toPT, value.ToString());
                }
                ptMatrixCW.AddRow(resultRow);
            }

            ptMatrixCW.SaveToDisk();

            string allPTNameFilePath = destFilePath + "_AllPTName.xlsx";
            Dictionary<string, int> allPTNameColumnDic = new Dictionary<string, int>();
            allPTNameColumnDic.Add("name", 0);
            allPTNameColumnDic.Add("count", 1);
            Dictionary<string, string> allPTNameColumnFormats = new Dictionary<string, string>();
            allPTNameColumnFormats.Add("count", "#0");
            ExcelWriter allPTNameEW = new ExcelWriter(allPTNameFilePath, "List", allPTNameColumnDic, allPTNameColumnFormats);
            for (int i = 0; i < allPTList.Count; i++)
            {
                string fromPT = allPTList[i];
                Dictionary<string, object> resultRow = new Dictionary<string, object>();
                resultRow.Add("name", fromPT);
                resultRow.Add("count", allPTCountDic[fromPT]);
                allPTNameEW.AddRow(resultRow);

            }
            allPTNameEW.SaveToDisk();

            string ptNameFilePath = destFilePath + "_PTName.xlsx";
            Dictionary<string, int> ptNameColumnDic = new Dictionary<string, int>();
            ptNameColumnDic.Add("name", 0);
            ExcelWriter ptNameEW = new ExcelWriter(ptNameFilePath, "List", ptNameColumnDic);
            for (int i = 0; i < ptList.Count; i++)
            {
                string fromPT = ptList[i];
                Dictionary<string, string> resultRow = new Dictionary<string, string>();
                resultRow.Add("name", fromPT);
                ptNameEW.AddRow(resultRow);

            }
            ptNameEW.SaveToDisk();


            string ptArrayFilePath = destFilePath + "_Array.txt";
            StringBuilder ptArrayStringBuilder = new StringBuilder();
            ptArrayStringBuilder.Append("arr = [");
            for (int i = 0; i < ptList.Count; i++)
            {
                string fromPT = ptList[i];
                ptArrayStringBuilder.Append((i == 0 ? "" : ", \r\n") + "[");
                Dictionary<string, string> resultRow = new Dictionary<string, string>();
                resultRow.Add("ptToPT", fromPT);
                Dictionary<string, int> ptDic = ptToPTDic.ContainsKey(fromPT) ? ptToPTDic[fromPT] : null;
                for (int j = 0; j < ptListDic.Count; j++)
                {
                    string toPT = ptList[j];
                    double value = fromPT == toPT ? 0 : (ptDic == null || !ptDic.ContainsKey(toPT) || ptDic[toPT] == 0 ? 2 * (double)maxTime : ((double)maxTime / (double)ptDic[toPT]));
                    resultRow.Add(toPT, value.ToString());
                    ptArrayStringBuilder.Append((j == 0 ? "" : ", ") + value.ToString());
                }
                ptMatrixCW.AddRow(resultRow);
                ptArrayStringBuilder.Append("]");
            }
            ptArrayStringBuilder.Append("]");
            FileHelper.SaveTextToFile(ptArrayStringBuilder.ToString(), ptArrayFilePath);
        } 
    }
}