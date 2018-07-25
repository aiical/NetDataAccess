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
using NetDataAccess.Base.CsvHelper;
using NetDataAccess.Base.Reader;

namespace NetDataAccess.Extended.GlassDoor
{
    public class GetRatingData : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        { 
            this.GetRatingInfos(listSheet);
            return true;
        }

        private ExcelWriter GetRaitingInfoExcelWriter(string destFilePath)
        { 

            Dictionary<string, int> columnDic = CommonUtil.InitStringIndexDic(new string[]{ 
                    "Company_Name", 
                    "Page_Company_Name",
                    "EmployerId",
                    "ItemName",
                    "ItemValue"});

            ExcelWriter ew = new ExcelWriter(destFilePath, "List", columnDic);
            return ew;
        }

        private void GetRatingInfos(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            string resultFilePath = Path.Combine(exportDir, "GlassDoor_RatingDetail.xlsx");

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;

            Dictionary<string, string> houseDic = new Dictionary<string, string>();

            ExcelWriter resultEW = this.GetRaitingInfoExcelWriter(resultFilePath);

            for (int i = 0; i < listSheet.RowCount; i++)
            {

                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string url = row[detailPageUrlColumnName];
                    string companyName = row["Company_Name"];
                    string pageCompanyName = row["Page_Company_Name"];
                    string employerId = row["EmployerId"];

                    string pageFilePath = this.RunPage.GetFilePath(url, pageSourceDir);
                    string jsonText = FileHelper.GetTextFromFile(pageFilePath);


                    try
                    {
                        JObject infoJo = JObject.Parse(jsonText);

                        JArray ratingArray = infoJo.GetValue("ratings") as JArray;

                        for (int j = 0; j < ratingArray.Count; j++)
                        {
                            JObject itemJo = ratingArray[j] as JObject;
                            string itemName = itemJo.GetValue("type").ToString();
                            string itemValue = itemJo.GetValue("value").ToString();

                            Dictionary<string, string> resultRow = new Dictionary<string, string>();
                            resultRow.Add("Company_Name", companyName);
                            resultRow.Add("Page_Company_Name", pageCompanyName);
                            resultRow.Add("EmployerId", employerId);
                            resultRow.Add("ItemName", itemName);
                            resultRow.Add("ItemValue", itemValue);
                            resultEW.AddRow(resultRow);
                        }
                    }
                    catch (Exception ex)
                    {
                        this.RunPage.InvokeAppendLogText(ex.Message + ", pageUrl = " + url, LogLevelType.System, true);
                        throw ex;
                    }
                }
            }

            resultEW.SaveToDisk();
        } 
    }
}