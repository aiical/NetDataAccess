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
                    "overallRating",
                    "ceoRating",
                    "bizOutlook",
                    "recommend",
                    "compAndBenefits",
                    "cultureAndValues",
                    "careerOpportunities",
                    "workLife",
                    "seniorManagement"
            });

            Dictionary<string, string> columnFormats = new Dictionary<string, string>();
            columnFormats.Add("overallRating", "#0.00");
            columnFormats.Add("ceoRating", "#0.00");
            columnFormats.Add("bizOutlook", "#0.00");
            columnFormats.Add("recommend", "#0.00");
            columnFormats.Add("compAndBenefits", "#0.00");
            columnFormats.Add("cultureAndValues", "#0.00");
            columnFormats.Add("careerOpportunities", "#0.00");
            columnFormats.Add("workLife", "#0.00");
            columnFormats.Add("seniorManagement", "#0.00");

            ExcelWriter ew = new ExcelWriter(destFilePath, "List", columnDic, columnFormats);
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
                        Nullable<decimal> overallRating = null;
                        Nullable<decimal> ceoRating = null;
                        Nullable<decimal> bizOutlook = null;
                        Nullable<decimal> recommend = null;
                        Nullable<decimal> compAndBenefits = null;
                        Nullable<decimal> cultureAndValues = null;
                        Nullable<decimal> careerOpportunities = null;
                        Nullable<decimal> workLife = null;
                        Nullable<decimal> seniorManagement = null;


                        for (int j = 0; j < ratingArray.Count; j++)
                        {
                            JObject itemJo = ratingArray[j] as JObject;
                            string itemName = itemJo.GetValue("type").ToString();
                            string itemValueStr = itemJo.GetValue("value").ToString();
                            Nullable<decimal> itemValue = itemValueStr.Length == 0 ? null : (Nullable<decimal>)decimal.Parse(itemValueStr);

                            switch (itemName)
                            {
                                case "overallRating":
                                    overallRating = itemValue;
                                    break;
                                case "ceoRating":
                                    ceoRating = itemValue;
                                    break;
                                case "bizOutlook":
                                    bizOutlook = itemValue;
                                    break;
                                case "recommend":
                                    recommend = itemValue;
                                    break;
                                case "compAndBenefits":
                                    compAndBenefits = itemValue;
                                    break;
                                case "cultureAndValues":
                                    cultureAndValues = itemValue;
                                    break;
                                case "careerOpportunities":
                                    careerOpportunities = itemValue;
                                    break;
                                case "workLife":
                                    workLife = itemValue;
                                    break;
                                case "seniorManagement":
                                    seniorManagement = itemValue;
                                    break;
                            }
                        }

                        Dictionary<string, object> resultRow = new Dictionary<string, object>();
                        resultRow.Add("Company_Name", companyName);
                        resultRow.Add("Page_Company_Name", pageCompanyName);
                        resultRow.Add("EmployerId", employerId);
                        resultRow.Add("overallRating", overallRating.HasValue ? (object)overallRating : null);
                        resultRow.Add("ceoRating", ceoRating.HasValue ? (object)ceoRating : null);
                        resultRow.Add("bizOutlook", bizOutlook.HasValue ? (object)bizOutlook : null);
                        resultRow.Add("recommend", recommend.HasValue ? (object)recommend : null);
                        resultRow.Add("compAndBenefits", compAndBenefits.HasValue ? (object)compAndBenefits : null);
                        resultRow.Add("cultureAndValues", cultureAndValues.HasValue ? (object)cultureAndValues : null);
                        resultRow.Add("careerOpportunities", careerOpportunities.HasValue ? (object)careerOpportunities : null);
                        resultRow.Add("workLife", workLife.HasValue ? (object)workLife : null);
                        resultRow.Add("seniorManagement", seniorManagement.HasValue ? (object)seniorManagement : null);
                        resultEW.AddRow(resultRow);
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