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

namespace NetDataAccess.Extended.Jiaoyu.Zhuanye
{
    public class GetBenkeMenleiListPage_jhcee_com : ExternalRunWebPage
    {
        public override byte[] GetRequestData_BeforeSendRequest(string pageUrl, Dictionary<string, string> listRow, Encoding encoding)
        {
            string id = listRow["id"];
            string data = "parentId=" + id;
            return encoding.GetBytes(data);
        }

        public override bool AfterAllGrab(IListSheet listSheet)
        {
            try
            {
                this.GetList(listSheet); 
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return true;
        }

        private void GetList(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("学科", 5);
            resultColumnDic.Add("学科id", 6); 
            resultColumnDic.Add("门类", 7);
            resultColumnDic.Add("门类id", 8); 
            string resultFilePath = Path.Combine(exportDir, "教育_本科_专业_jhcee_com.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null); 
            
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"];
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string localFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir);

                    try
                    {
                        string pageFileText = FileHelper.GetTextFromFile(localFilePath);
                        JObject rootJo = JObject.Parse(pageFileText);

                        JArray itemJsons = rootJo["data"] as JArray;
                        foreach (JObject itemJson in itemJsons)
                        {
                            string name = itemJson["name"].ToString();
                            string id = itemJson["id"].ToString();
                            string parentId = itemJson["parentId"].ToString();

                            Dictionary<string, string> f2vs = new Dictionary<string, string>();
                            f2vs.Add("detailPageUrl", "http://www.jhcee.com/specialized/loadByParentId.json?parentId=" + id);
                            f2vs.Add("detailPageName", id);
                            f2vs.Add("门类", name);
                            f2vs.Add("门类id", id);
                            f2vs.Add("学科", row["name"]);
                            f2vs.Add("学科id", row["id"]);
                            resultEW.AddRow(f2vs);
                        } 
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    } 
                }
            }
            resultEW.SaveToDisk();
        }
    }
}