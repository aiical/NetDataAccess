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
    public class ProcessBaiKeItemLinkageResult : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            this.ProcessLinkageResult(listSheet);
            return true;
        }

        private void ProcessLinkageResult(IListSheet listSheet)
        {
            try
            {
                string[] parameters = this.Parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                string sourceFilePath = parameters[0];
                string nameFilePath = parameters[1];
                string destFilePath = parameters[2];

                ExcelReader nameER = new ExcelReader(nameFilePath);
                int nameRowCount = nameER.GetRowCount();
                Dictionary<string, string> nameDic = new Dictionary<string, string>();
                for (int i = 0; i < nameRowCount; i++)
                {
                    Dictionary<string, string> nameRow = nameER.GetFieldValues(i);
                    nameDic.Add(i.ToString(), nameRow["name"]);
                }

                ExcelReader er = new ExcelReader(sourceFilePath);
                int sourceRowCount = er.GetRowCount();

                List<string> tagList = new List<string>();
                int parentNodeIndex = sourceRowCount;
                Dictionary<string, string[]> nodeDic = new Dictionary<string, string[]>();
                for (int i = 0; i < sourceRowCount; i++)
                {
                    parentNodeIndex++;
                    Dictionary<string, string> sourceRow = er.GetFieldValues(i);
                    nodeDic.Add(parentNodeIndex.ToString(), new string[] { sourceRow["0"], sourceRow["1"] });
                }

                JObject json = new JObject();
                this.AddToJson(json, parentNodeIndex.ToString(), nodeDic, nameDic);
                FileHelper.SaveTextToFile(json.ToString(), destFilePath);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void AddToJson(JObject json, string indexStr, Dictionary<string, string[]> nodeDic, Dictionary<string, string> nameDic)
        {
            string name = indexStr + (!nameDic.ContainsKey(indexStr) ? "" : ("." + nameDic[indexStr]));
            json.Add("name", name);
            json.Add("collapsed", false);

            if (nodeDic.ContainsKey(indexStr))
            {
                JArray childrenArray = new JArray();

                string[] nextLevelNodes = nodeDic[indexStr];
                JObject childJsonA = new JObject();
                childrenArray.Add(childJsonA);
                string indexA = nextLevelNodes[0];
                this.AddToJson(childJsonA, indexA, nodeDic, nameDic);


                JObject childJsonB = new JObject();
                childrenArray.Add(childJsonB);
                string indexB = nextLevelNodes[1];
                this.AddToJson(childJsonB, indexB, nodeDic, nameDic);


                json.Add("children", childrenArray);
            } 
        }
    }
}