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
    public class GetBaiKeTagRelateMatrix : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            this.GetTagsMatrix(listSheet);
            return true;
        }

        private void GetTagsMatrix(IListSheet listSheet)
        {
            string[] parameters = this.Parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            string sourceFilePath = parameters[0];
            string destFilePath = parameters[1];

            ExcelReader er = new ExcelReader(sourceFilePath);
            int sourceRowCount = er.GetRowCount();

            List<string> tagList = new List<string>();
            for (int i = 0; i < sourceRowCount; i++)
            {
                Dictionary<string,string> sourceRow = er.GetFieldValues(i);
                string[] itemTags = sourceRow["tags"].Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string itemTag in itemTags)
                {
                    if (!tagList.Contains(itemTag))
                    {
                        tagList.Add(itemTag);
                    }
                }
            }

            int maxTime = 1;

            Dictionary<string, Dictionary<string, int>> tagToTagDic = new Dictionary<string, Dictionary<string, int>>();
            for (int i = 0; i < sourceRowCount; i++)
            {
                Dictionary<string, string> sourceRow = er.GetFieldValues(i); 
                string[] itemTags = sourceRow["tags"].Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string fromTag in itemTags)
                {
                    if (!tagToTagDic.ContainsKey(fromTag))
                    {
                        tagToTagDic.Add(fromTag, new Dictionary<string, int>());
                    }
                    Dictionary<string, int> tagDic = tagToTagDic[fromTag];

                    if (!tagDic.ContainsKey(fromTag))
                    {
                        tagDic.Add(fromTag, 1);
                    }
                    else
                    {
                        tagDic[fromTag] = tagDic[fromTag] + 1;
                    }

                    foreach (string toTag in itemTags)
                    {
                        if (fromTag != toTag)
                        {
                            if (!tagDic.ContainsKey(toTag))
                            {
                                tagDic.Add(toTag, 1);
                            }
                            else
                            {
                                int tmpValue = tagDic[toTag] + 1;
                                tagDic[toTag] = tmpValue;
                                if (tmpValue > maxTime)
                                {
                                    maxTime = tmpValue;
                                }
                            }
                        }
                    }
                }
            }


            
            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("tagToTag", 0); 
            for (int i = 0; i < tagList.Count; i++)
            {
                resultColumnDic.Add(tagList[i], i + 1);
            }
            
            CsvWriter tagMatrixCW = new CsvWriter(destFilePath, resultColumnDic);

            foreach (string fromTag in tagList)
            { 

                Dictionary<string, string> resultRow = new Dictionary<string,string>();
                resultRow.Add("tagToTag", fromTag);
                Dictionary<string, int> tagDic = tagToTagDic.ContainsKey(fromTag) ? tagToTagDic[fromTag] : null;
                foreach (string toTag in tagList)
                {
                    double value = fromTag == toTag ? 0 : (tagDic == null || !tagDic.ContainsKey(toTag) || tagDic[toTag] == 0 ? 2 * (double)maxTime : ((double)maxTime / (double)tagDic[toTag]));
                    resultRow.Add(toTag, value.ToString());
                }
                tagMatrixCW.AddRow(resultRow);
            }

            tagMatrixCW.SaveToDisk();

            string tagNameFilePath = destFilePath + "_TagName.xlsx";
            Dictionary<string, int> tagNameColumnDic = new Dictionary<string, int>();
            tagNameColumnDic.Add("name", 0);
            ExcelWriter tagNameEW = new ExcelWriter(tagNameFilePath, "List", tagNameColumnDic);
            for (int i = 0; i < tagList.Count; i++)
            {
                string fromTag = tagList[i];
                Dictionary<string, string> resultRow = new Dictionary<string, string>();
                resultRow.Add("name", fromTag);
                tagNameEW.AddRow(resultRow);

            }
            tagNameEW.SaveToDisk();


            string tagArrayFilePath = destFilePath + "_Array.txt";
            StringBuilder tagArrayStringBuilder = new StringBuilder();
            tagArrayStringBuilder.Append("arr = [");
            for (int i = 0; i < tagList.Count; i++)
            {
                string fromTag = tagList[i];
                tagArrayStringBuilder.Append((i == 0 ? "" : ", \r\n") + "[");
                Dictionary<string, string> resultRow = new Dictionary<string, string>();
                resultRow.Add("tagToTag", fromTag);
                Dictionary<string, int> tagDic = tagToTagDic.ContainsKey(fromTag) ? tagToTagDic[fromTag] : null;
                for (int j = 0; j < tagList.Count; j++)
                {
                    string toTag = tagList[j];
                    double value = fromTag == toTag ? 0 : (tagDic == null || !tagDic.ContainsKey(toTag) || tagDic[toTag] == 0 ? 2 * (double)maxTime : ((double)maxTime / (double)tagDic[toTag]));
                    resultRow.Add(toTag, value.ToString());
                    tagArrayStringBuilder.Append((j == 0 ? "" : ", ") + value.ToString());
                }
                tagMatrixCW.AddRow(resultRow);
                tagArrayStringBuilder.Append("]");
            }
            tagArrayStringBuilder.Append("]");
            FileHelper.SaveTextToFile(tagArrayStringBuilder.ToString(), tagArrayFilePath);
        } 
    }
}