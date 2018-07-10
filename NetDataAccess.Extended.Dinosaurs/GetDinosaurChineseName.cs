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
using System.Web;

namespace NetDataAccess.Extended.Dinosaurs
{
    public class GetDinosaurChineseName : ExternalRunWebPage
    {
        public override void CheckRequestCompleteFile(string webPageText, Dictionary<string, string> listRow)
        {
            JObject rootJo = JObject.Parse(webPageText);
            JToken error_codeToken = rootJo.SelectToken("error_code");
            JToken error_msgToken = rootJo.SelectToken("error_msg");
            if (error_codeToken != null)
            {
                throw new Exception("error_code = " + error_codeToken.ToString() + ", " + "error_msg = " + error_msgToken.ToString());
            }
        }

        public override bool AfterAllGrab(IListSheet listSheet)
        {
            this.GetChineseName(listSheet);
            return true;
        }

        private void GetChineseName(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("name", 0);
            resultColumnDic.Add("name_cn", 1);
            resultColumnDic.Add("length", 2);
            resultColumnDic.Add("description", 3);
            resultColumnDic.Add("diet", 4);
            resultColumnDic.Add("country", 5);
            resultColumnDic.Add("period", 6);
            resultColumnDic.Add("teeth", 7);
            resultColumnDic.Add("how_it_moved", 8);
            resultColumnDic.Add("food", 9);
            resultColumnDic.Add("content", 10);
            resultColumnDic.Add("taxonomy", 11);
            resultColumnDic.Add("type_species", 12);
            resultColumnDic.Add("url", 13); 
            string resultFilePath = Path.Combine(exportDir, "www.nhm.ac.uk恐龙信息(中英文).xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            Dictionary<string, string> urlDic = new Dictionary<string, string>();
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"];
                string name = row["name"];
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                { 
                    string localFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir);
                    TextReader tr = null;

                    try
                    {
                        tr = new StreamReader(localFilePath, Encoding.UTF8);
                        string js = tr.ReadToEnd();

                        JObject rootJo = JObject.Parse(js);
                        JArray trListJsons = rootJo.SelectToken("trans_result") as JArray;
                        if (trListJsons.Count != 0)
                        {
                            for (int j = 0; j < trListJsons.Count; j++)
                            {
                                JObject trJson = trListJsons[j] as JObject;
                                if (name == trJson.SelectToken("src").ToString())
                                {
                                    string name_cn = HttpUtility.HtmlDecode(trJson.SelectToken("dst").ToString());
                                    Dictionary<string, string> f2vs = new Dictionary<string, string>();
                                    f2vs.Add("name", name);
                                    f2vs.Add("name_cn", name_cn);
                                    f2vs.Add("length", row["length"]);
                                    f2vs.Add("description", row["description"]);
                                    f2vs.Add("diet", row["diet"]);
                                    f2vs.Add("country", row["country"]);
                                    f2vs.Add("period", row["period"]);
                                    f2vs.Add("teeth", row["teeth"]);
                                    f2vs.Add("how_it_moved", row["how_it_moved"]);
                                    f2vs.Add("food", row["food"]);
                                    f2vs.Add("content", row["content"]);
                                    f2vs.Add("taxonomy", row["taxonomy"]);
                                    f2vs.Add("type_species", row["type_species"]);
                                    f2vs.Add("url", detailUrl);
                                    resultEW.AddRow(f2vs);
                                    break;
                                }
                            }
                        } 
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                    finally
                    {
                        if (tr != null)
                        {
                            tr.Close();
                            tr.Dispose();
                        }
                    }
                }
            }
            resultEW.SaveToDisk();
        }
         
    }
}