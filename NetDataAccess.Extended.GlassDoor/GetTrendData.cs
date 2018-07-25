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
    public class GetTrendData : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            this.GetTrendInfos(listSheet);
            return true;
        }

        private ExcelWriter GetTrendInfoExcelWriter(string destFilePath)
        { 

            Dictionary<string, int> columnDic = CommonUtil.InitStringIndexDic(new string[]{ 
                    "Company_Name", 
                    "Page_Company_Name",
                    "EmployerId",
                    "Date",
                    "EmployerRating"});

            ExcelWriter ew = new ExcelWriter(destFilePath, "List", columnDic);
            return ew;
        }

        private void GetTrendInfos(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            string resultFilePath = Path.Combine(exportDir, "GlassDoor_TrendDetail.xlsx");

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;

            Dictionary<string, string> houseDic = new Dictionary<string, string>();

            ExcelWriter resultEW = this.GetTrendInfoExcelWriter(resultFilePath);

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

                        JArray dateArray = infoJo.GetValue("dates") as JArray;
                        JArray employerRatingArray = infoJo.GetValue("employerRatings") as JArray;

                        for (int j = 0; j < dateArray.Count; j++)
                        {
                            string date = dateArray[j].ToString();
                            string employerRating = employerRatingArray[j].ToString(); 

                            Dictionary<string, string> resultRow = new Dictionary<string, string>();
                            resultRow.Add("Company_Name", companyName);
                            resultRow.Add("Page_Company_Name", pageCompanyName);
                            resultRow.Add("EmployerId", employerId);
                            resultRow.Add("Date", date);
                            resultRow.Add("EmployerRating", employerRating);
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