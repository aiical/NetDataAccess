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
    public class GetBenkeZhuanyeListPage_jhcee_com : ExternalRunWebPage
    {
        public override byte[] GetRequestData_BeforeSendRequest(string pageUrl, Dictionary<string, string> listRow, Encoding encoding)
        {
            string id = listRow["门类id"];
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
            resultColumnDic.Add("学科", 0);
            resultColumnDic.Add("学科id", 1); 
            resultColumnDic.Add("门类", 2);
            resultColumnDic.Add("门类id", 3);
            resultColumnDic.Add("专业", 4);
            resultColumnDic.Add("专业id", 5);
            resultColumnDic.Add("培养目标", 6);
            resultColumnDic.Add("培养要求", 7);
            resultColumnDic.Add("毕业生应获得以下几方面的知识和能力", 8);
            resultColumnDic.Add("主要课程", 9);
            resultColumnDic.Add("主要实践性教学环节", 10);
            resultColumnDic.Add("修业年薪", 11);
            resultColumnDic.Add("授予学位", 12);
            resultColumnDic.Add("主干学科", 13);
            resultColumnDic.Add("url", 14); 
            string resultFilePath = Path.Combine(exportDir, "教育_本科_专业_详情_jhcee_com.xlsx");
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
                            f2vs.Add("门类", row["门类"]);
                            f2vs.Add("门类id", row["门类id"]);
                            f2vs.Add("学科", row["学科"]);
                            f2vs.Add("学科id", row["学科id"]);
                            f2vs.Add("专业", name);
                            f2vs.Add("专业id", id);
                            f2vs.Add("培养目标", itemJson["trainingObjective"].ToString());
                            f2vs.Add("培养要求", itemJson["trainingRequest"].ToString());
                            f2vs.Add("毕业生应获得以下几方面的知识和能力", itemJson["graduatesSkills"].ToString());
                            f2vs.Add("主要课程", itemJson["specializedCourses"].ToString());
                            f2vs.Add("主要实践性教学环节", itemJson["practiceCourse"].ToString());
                            f2vs.Add("修业年薪", itemJson["educationalSystem"].ToString());
                            f2vs.Add("授予学位", itemJson["degreesConferred"].ToString());
                            f2vs.Add("主干学科", itemJson["majorDisciplines"].ToString());
                            f2vs.Add("url", "http://www.jhcee.com/chooseSpecialtyDetails.htm?id=" + id); 
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