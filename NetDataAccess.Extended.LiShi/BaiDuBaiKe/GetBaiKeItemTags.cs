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
    public class GetBaiKeItemTags : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        { 
            this.GetItemTagsTypes(listSheet);
            return true;
        }

        private void GetItemTagsTypes(IListSheet listSheet)
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
            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("url", 0);
            resultColumnDic.Add("itemId", 1);
            resultColumnDic.Add("itemName", 2);
            resultColumnDic.Add("tags", 3);
            for (int i = 0; i < tagList.Count; i++)
            {
                resultColumnDic.Add(tagList[i], i + 4);
            }


            CsvWriter itemTagMatrixCW = new CsvWriter(destFilePath, resultColumnDic);

            for (int i = 0; i < sourceRowCount; i++)
            {
                Dictionary<string, string> sourceRow = er.GetFieldValues(i);

                Dictionary<string, string> resultRow = new Dictionary<string,string>();
                resultRow.Add("url", sourceRow["url"]);
                resultRow.Add("itemId",  sourceRow["itemId"]);
                resultRow.Add("itemName",  sourceRow["itemName"]);
                resultRow.Add("tags",  sourceRow["tags"]);
                string tagsStr = sourceRow["tags"];
                string[] itemTags = sourceRow["tags"].Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);

                foreach (string itemTag in itemTags)
                {
                    resultRow.Add(itemTag, "1");
                }
                itemTagMatrixCW.AddRow(resultRow);
            }

            itemTagMatrixCW.SaveToDisk();
        } 
    }
}