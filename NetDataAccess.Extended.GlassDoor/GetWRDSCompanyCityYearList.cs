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
using NPOI.XSSF.UserModel;

namespace NetDataAccess.Extended.GlassDoor
{
    public class GetWRDSCompanyCityYearList : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        {

            string[] parameters = this.Parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            string sourceFilePath = parameters[0]; 
            string destFilePath = parameters[1];

            try
            {
                this.ProcessWRDSCompanyCityYearList(sourceFilePath, destFilePath);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return true;
        }

        private ExcelWriter GetExcelWriter(string destFilePath)
        {

            Dictionary<string, int> columnDic = CommonUtil.InitStringIndexDic(new string[]{ 
                    "Auditor fkey",
                    "Auditor name",
                    "Source company name",
                    "Search company name",
                    "GlassDoor company name",
                    "Auditor city",
                    "Auditor state name",
                    "Auditor state",
                    "Auditor country",
                    "Auditor region",
                    "City_State",
                    "Year",
                    "CompanyInWRDS",
                    "CompanyCityInWRDS",
                    "Count_position_directors_managers",
                    "Count_position_others",
                    "Count_FullTime Job",
                    "Count_PartTime Job",
                    "Count_Ohters Job",
                    "Average_Words_Pros",
                    "Average_Words_Cons",
                    "Total_CountNumber_Rating",
                    "Total_Average_Rating",
                    "Total_Average_WorkLifeBalance",
                    "Total_Average_CultureValues",
                    "Total_Average_CareerOpportunities",
                    "Total_Average_CompBenefits",
                    "Total_Average_SeniorManagement",
                    "Total_CountNumber_Recommends",
                    "Total_CountNumber_Positive Outlook",
                    "Total_CountNumber_Negative Outlook",
                    "Total_CountNumber_Neutral Outlook",
                    "Total_CountNumber_NULL Outlook",
                    "Total_CountNumber_No opinion of CEO",
                    "Total_CountNumber_Approves of CEO",
                    "Total_CountNumber_Disapproves of CEO",
                    "Total_CountNumber_NULL value approves_CEO",
                    "Former Employee_CountNumber_Rating",
                    "Former Employee_Average_Rating",
                    "Former Employee_Average_WorkLifeBalance",
                    "Former Employee_Average_CultureValues",
                    "Former Employee_Average_CareerOpportunities",
                    "Former Employee_Average_CompBenefits",
                    "Former Employee_Average_SeniorManagement",
                    "Former Employee_CountNumber_Recommends",
                    "Former Employee_CountNumber_Positive Outlook",
                    "Former Employee_CountNumber_Negative Outlook",
                    "Former Employee_CountNumber_Neutral Outlook",
                    "Former Employee_CountNumber_NULL Outlook",
                    "Former Employee_CountNumber_No opinion of CEO",
                    "Former Employee_CountNumber_Approves of CEO",
                    "Former Employee_CountNumber_Disapproves of CEO",
                    "Former Employee_CountNumber_NULL value approves_CEO",
                    "Current Employee_CountNumber_Rating",
                    "Current Employee_Average_Rating",
                    "Current Employee_Average_WorkLifeBalance",
                    "Current Employee_Average_CultureValues",
                    "Current Employee_Average_CareerOpportunities",
                    "Current Employee_Average_CompBenefits",
                    "Current Employee_Average_SeniorManagement",
                    "Current Employee_CountNumber_Recommends",
                    "Current Employee_CountNumber_Positive Outlook",
                    "Current Employee_CountNumber_Negative Outlook",
                    "Current Employee_CountNumber_Neutral Outlook",
                    "Current Employee_CountNumber_NULL Outlook",
                    "Current Employee_CountNumber_No opinion of CEO",
                    "Current Employee_CountNumber_Approves of CEO",
                    "Current Employee_CountNumber_Disapproves of CEO",
                    "Current Employee_CountNumber_NULL value approves_CEO",
                    "Unknown Employee_CountNumber_Rating",
                    "Unknown Employee_Average_Rating",
                    "Unknown Employee_Average_WorkLifeBalance",
                    "Unknown Employee_Average_CultureValues",
                    "Unknown Employee_Average_CareerOpportunities",
                    "Unknown Employee_Average_CompBenefits",
                    "Unknown Employee_Average_SeniorManagement",
                    "Unknown Employee_CountNumber_Recommends",
                    "Unknown Employee_CountNumber_Positive Outlook",
                    "Unknown Employee_CountNumber_Negative Outlook",
                    "Unknown Employee_CountNumber_Neutral Outlook",
                    "Unknown Employee_CountNumber_NULL Outlook",
                    "Unknown Employee_CountNumber_No opinion of CEO",
                    "Unknown Employee_CountNumber_Approves of CEO",
                    "Unknown Employee_CountNumber_Disapproves of CEO",
                    "Unknown Employee_CountNumber_NULL value approves_CEO"
                    });
            Dictionary<string, string> columnFormats = new Dictionary<string, string>();
            columnFormats.Add("Year", "#0");
            columnFormats.Add("Count_position_directors_managers", "#0");
            columnFormats.Add("Count_position_others", "#0");
            columnFormats.Add("Count_FullTime Job", "#0");
            columnFormats.Add("Count_PartTime Job", "#0");
            columnFormats.Add("Count_Ohters Job", "#0");
            columnFormats.Add("Average_Words_Pros", "#0.0000");
            columnFormats.Add("Average_Words_Cons", "#0.0000");
            columnFormats.Add("Total_CountNumber_Rating", "#0");
            columnFormats.Add("Total_Average_Rating", "#0.0000");
            columnFormats.Add("Total_Average_WorkLifeBalance", "#0.0000");
            columnFormats.Add("Total_Average_CultureValues", "#0.0000");
            columnFormats.Add("Total_Average_CareerOpportunities", "#0.0000");
            columnFormats.Add("Total_Average_CompBenefits", "#0.0000");
            columnFormats.Add("Total_Average_SeniorManagement", "#0.0000");
            columnFormats.Add("Total_CountNumber_Recommends", "#0");
            columnFormats.Add("Total_CountNumber_Positive Outlook", "#0");
            columnFormats.Add("Total_CountNumber_Negative Outlook", "#0");
            columnFormats.Add("Total_CountNumber_Neutral Outlook", "#0");
            columnFormats.Add("Total_CountNumber_NULL Outlook", "#0");
            columnFormats.Add("Total_CountNumber_No opinion of CEO", "#0");
            columnFormats.Add("Total_CountNumber_Approves of CEO", "#0");
            columnFormats.Add("Total_CountNumber_Disapproves of CEO", "#0");
            columnFormats.Add("Total_CountNumber_NULL value approves_CEO", "#0");
            columnFormats.Add("Former Employee_CountNumber_Rating", "#0");
            columnFormats.Add("Former Employee_Average_Rating", "#0.0000");
            columnFormats.Add("Former Employee_Average_WorkLifeBalance", "#0.0000");
            columnFormats.Add("Former Employee_Average_CultureValues", "#0.0000");
            columnFormats.Add("Former Employee_Average_CareerOpportunities", "#0.0000");
            columnFormats.Add("Former Employee_Average_CompBenefits", "#0.0000");
            columnFormats.Add("Former Employee_Average_SeniorManagement", "#0.0000");
            columnFormats.Add("Former Employee_CountNumber_Recommends", "#0");
            columnFormats.Add("Former Employee_CountNumber_Positive Outlook", "#0");
            columnFormats.Add("Former Employee_CountNumber_Negative Outlook", "#0");
            columnFormats.Add("Former Employee_CountNumber_Neutral Outlook", "#0");
            columnFormats.Add("Former Employee_CountNumber_NULL Outlook", "#0");
            columnFormats.Add("Former Employee_CountNumber_No opinion of CEO", "#0");
            columnFormats.Add("Former Employee_CountNumber_Approves of CEO", "#0");
            columnFormats.Add("Former Employee_CountNumber_Disapproves of CEO", "#0");
            columnFormats.Add("Former Employee_CountNumber_NULL value approves_CEO", "#0");
            columnFormats.Add("Current Employee_CountNumber_Rating", "#0");
            columnFormats.Add("Current Employee_Average_Rating", "#0.0000");
            columnFormats.Add("Current Employee_Average_WorkLifeBalance", "#0.0000");
            columnFormats.Add("Current Employee_Average_CultureValues", "#0.0000");
            columnFormats.Add("Current Employee_Average_CareerOpportunities", "#0.0000");
            columnFormats.Add("Current Employee_Average_CompBenefits", "#0.0000");
            columnFormats.Add("Current Employee_Average_SeniorManagement", "#0.0000");
            columnFormats.Add("Current Employee_CountNumber_Recommends", "#0");
            columnFormats.Add("Current Employee_CountNumber_Positive Outlook", "#0");
            columnFormats.Add("Current Employee_CountNumber_Negative Outlook", "#0");
            columnFormats.Add("Current Employee_CountNumber_Neutral Outlook", "#0");
            columnFormats.Add("Current Employee_CountNumber_NULL Outlook", "#0");
            columnFormats.Add("Current Employee_CountNumber_No opinion of CEO", "#0");
            columnFormats.Add("Current Employee_CountNumber_Approves of CEO", "#0");
            columnFormats.Add("Current Employee_CountNumber_Disapproves of CEO", "#0");
            columnFormats.Add("Current Employee_CountNumber_NULL value approves_CEO", "#0");
            columnFormats.Add("Unknown Employee_CountNumber_Rating", "#0");
            columnFormats.Add("Unknown Employee_Average_Rating", "#0.0000");
            columnFormats.Add("Unknown Employee_Average_WorkLifeBalance", "#0.0000");
            columnFormats.Add("Unknown Employee_Average_CultureValues", "#0.0000");
            columnFormats.Add("Unknown Employee_Average_CareerOpportunities", "#0.0000");
            columnFormats.Add("Unknown Employee_Average_CompBenefits", "#0.0000");
            columnFormats.Add("Unknown Employee_Average_SeniorManagement", "#0.0000");
            columnFormats.Add("Unknown Employee_CountNumber_Recommends", "#0");
            columnFormats.Add("Unknown Employee_CountNumber_Positive Outlook", "#0");
            columnFormats.Add("Unknown Employee_CountNumber_Negative Outlook", "#0");
            columnFormats.Add("Unknown Employee_CountNumber_Neutral Outlook", "#0");
            columnFormats.Add("Unknown Employee_CountNumber_NULL Outlook", "#0");
            columnFormats.Add("Unknown Employee_CountNumber_No opinion of CEO", "#0");
            columnFormats.Add("Unknown Employee_CountNumber_Approves of CEO", "#0");
            columnFormats.Add("Unknown Employee_CountNumber_Disapproves of CEO", "#0");
            columnFormats.Add("Unknown Employee_CountNumber_NULL value approves_CEO", "#0");

            ExcelWriter ew = new ExcelWriter(destFilePath, "List", columnDic, columnFormats);
            return ew;
        }

        private Dictionary<string, object> ProcessWRDSCompanyCityYearRow(string key, string company, string city, string year, List<Dictionary<string, string>> sourceRows)
        {
            Dictionary<string, object> resultRow = new Dictionary<string, object>();
            string vCompany_name = company;
            string vCity = city;
            decimal vYear = decimal.Parse(year);
            decimal vCount_position_directors_managers = this.GetCountValue(sourceRows, "Position_directors_managers", "1");
            decimal vCount_position_others = this.GetCountValue(sourceRows, "Position_directors_managers", "0");
            decimal vCount_FullTime_Job = this.GetCountValue(sourceRows, "Full&Part&other-Time Job", "2");
            decimal vCount_PartTime_Job = this.GetCountValue(sourceRows, "Full&Part&other-Time Job", "1");
            decimal vCount_Ohters_Job = this.GetCountValue(sourceRows, "Full&Part&other-Time Job", "0");
            decimal vAverage_Words_Pros = this.GetAverageValue(sourceRows, "Words_Pros");
            decimal vAverage_Words_Cons = this.GetAverageValue(sourceRows, "Words_Cons");

            decimal vTotal_CountNumber_Rating = this.GetCountValueHasValue(sourceRows, "Rating");
            decimal vTotal_Average_Rating = this.GetAverageHasValue(sourceRows, "Rating");
            decimal vTotal_Average_WorkLifeBalance = this.GetAverageHasValue(sourceRows, "WorkLifeBalance");
            decimal vTotal_Average_CultureValues = this.GetAverageHasValue(sourceRows, "CultureValues");
            decimal vTotal_Average_CareerOpportunities = this.GetAverageHasValue(sourceRows, "CareerOpportunities");
            decimal vTotal_Average_CompBenefits = this.GetAverageHasValue(sourceRows, "CompBenefits");
            decimal vTotal_Average_SeniorManagement = this.GetAverageHasValue(sourceRows, "SeniorManagement");
            decimal vTotal_CountNumber_Recommends = this.GetCountValue(sourceRows, "Recommends", "Recommends");
            decimal vTotal_CountNumber_Positive_Outlook = this.GetCountValue(sourceRows, "Outlook", "Positive Outlook");
            decimal vTotal_CountNumber_Negative_Outlook = this.GetCountValue(sourceRows, "Outlook", "Negative Outlook");
            decimal vTotal_CountNumber_Neutral_Outlook = this.GetCountValue(sourceRows, "Outlook", "Neutral Outlook");
            decimal vTotal_CountNumber_NULL_Outlook = this.GetCountValueNullValue(sourceRows, "Outlook");
            decimal vTotal_CountNumber_No_opinion_of_CEO = this.GetCountValue(sourceRows, "OptionOfCEO", "No opinion of CEO");
            decimal vTotal_CountNumber_Approves_of_CEO = this.GetCountValue(sourceRows, "OptionOfCEO", "Approves of CEO");
            decimal vTotal_CountNumber_Disapproves_of_CEO = this.GetCountValue(sourceRows, "OptionOfCEO", "Disapproves of CEO");
            decimal vTotal_CountNumber_NULL_value_approves_CEO = this.GetCountValueNullValue(sourceRows, "OptionOfCEO");

            decimal vFormer_Employee_CountNumber_Rating = this.GetCountValueHasValue(sourceRows, "Employee", "Former Employee", "Rating");
            decimal vFormer_Employee_Average_Rating = this.GetAverageHasValue(sourceRows, "Employee", "Former Employee", "Rating");
            decimal vFormer_Employee_Average_WorkLifeBalance = this.GetAverageHasValue(sourceRows, "Employee", "Former Employee", "WorkLifeBalance");
            decimal vFormer_Employee_Average_CultureValues = this.GetAverageHasValue(sourceRows, "Employee", "Former Employee", "CultureValues");
            decimal vFormer_Employee_Average_CareerOpportunities = this.GetAverageHasValue(sourceRows, "Employee", "Former Employee", "CareerOpportunities");
            decimal vFormer_Employee_Average_CompBenefits = this.GetAverageHasValue(sourceRows, "Employee", "Former Employee", "CompBenefits");
            decimal vFormer_Employee_Average_SeniorManagement = this.GetAverageHasValue(sourceRows, "Employee", "Former Employee", "SeniorManagement");
            decimal vFormer_Employee_CountNumber_Recommends = this.GetCountValue(sourceRows, "Employee", "Former Employee", "Recommends", "Recommends");
            decimal vFormer_Employee_CountNumber_Positive_Outlook = this.GetCountValue(sourceRows, "Employee", "Former Employee", "Outlook", "Positive Outlook");
            decimal vFormer_Employee_CountNumber_Negative_Outlook = this.GetCountValue(sourceRows, "Employee", "Former Employee", "Outlook", "Negative Outlook");
            decimal vFormer_Employee_CountNumber_Neutral_Outlook = this.GetCountValue(sourceRows, "Employee", "Former Employee", "Outlook", "Neutral Outlook");
            decimal vFormer_Employee_CountNumber_NULL_Outlook = this.GetCountValueNullValue(sourceRows, "Employee", "Former Employee", "Outlook");
            decimal vFormer_Employee_CountNumber_No_opinion_of_CEO = this.GetCountValue(sourceRows, "Employee", "Former Employee", "OptionOfCEO", "No opinion of CEO");
            decimal vFormer_Employee_CountNumber_Approves_of_CEO = this.GetCountValue(sourceRows, "Employee", "Former Employee", "OptionOfCEO", "Approves of CEO");
            decimal vFormer_Employee_CountNumber_Disapproves_of_CEO = this.GetCountValue(sourceRows, "Employee", "Former Employee", "OptionOfCEO", "Disapproves of CEO");
            decimal vFormer_Employee_CountNumber_NULL_value_approves_CEO = this.GetCountValueNullValue(sourceRows, "Employee", "Former Employee", "OptionOfCEO");

            decimal vCurrent_Employee_CountNumber_Rating = this.GetCountValueHasValue(sourceRows, "Employee", "Current Employee", "Rating");
            decimal vCurrent_Employee_Average_Rating = this.GetAverageHasValue(sourceRows, "Employee", "Current Employee", "Rating");
            decimal vCurrent_Employee_Average_WorkLifeBalance = this.GetAverageHasValue(sourceRows, "Employee", "Current Employee", "WorkLifeBalance");
            decimal vCurrent_Employee_Average_CultureValues =this.GetAverageHasValue(sourceRows, "Employee", "Current Employee", "CultureValues");
            decimal vCurrent_Employee_Average_CareerOpportunities = this.GetAverageHasValue(sourceRows, "Employee", "Current Employee", "CareerOpportunities");
            decimal vCurrent_Employee_Average_CompBenefits = this.GetAverageHasValue(sourceRows, "Employee", "Current Employee", "CompBenefits");
            decimal vCurrent_Employee_Average_SeniorManagement = this.GetAverageHasValue(sourceRows, "Employee", "Current Employee", "SeniorManagement");
            decimal vCurrent_Employee_CountNumber_Recommends =  this.GetCountValue(sourceRows, "Employee", "Current Employee", "Recommends", "Recommends");
            decimal vCurrent_Employee_CountNumber_Positive_Outlook = this.GetCountValue(sourceRows, "Employee", "Current Employee", "Outlook", "Positive Outlook");
            decimal vCurrent_Employee_CountNumber_Negative_Outlook = this.GetCountValue(sourceRows, "Employee", "Current Employee", "Outlook", "Negative Outlook");
            decimal vCurrent_Employee_CountNumber_Neutral_Outlook =this.GetCountValue(sourceRows, "Employee", "Current Employee", "Outlook", "Neutral Outlook");
            decimal vCurrent_Employee_CountNumber_NULL_Outlook = this.GetCountValueNullValue(sourceRows, "Employee", "Current Employee", "Outlook");
            decimal vCurrent_Employee_CountNumber_No_opinion_of_CEO = this.GetCountValue(sourceRows, "Employee", "Current Employee", "OptionOfCEO", "No opinion of CEO");
            decimal vCurrent_Employee_CountNumber_Approves_of_CEO = this.GetCountValue(sourceRows, "Employee", "Current Employee", "OptionOfCEO", "Approves of CEO");
            decimal vCurrent_Employee_CountNumber_Disapproves_of_CEO = this.GetCountValue(sourceRows, "Employee", "Current Employee", "OptionOfCEO", "Disapproves of CEO");
            decimal vCurrent_Employee_CountNumber_NULL_value_approves_CEO = this.GetCountValueNullValue(sourceRows, "Employee", "Current Employee", "OptionOfCEO");

            decimal vUnknown_Employee_CountNumber_Rating = this.GetCountValueHasValue(sourceRows, "Employee", "", "Rating");
            decimal vUnknown_Employee_Average_Rating = this.GetAverageHasValue(sourceRows, "Employee", "", "Rating");
            decimal vUnknown_Employee_Average_WorkLifeBalance = this.GetAverageHasValue(sourceRows, "Employee", "", "WorkLifeBalance");
            decimal vUnknown_Employee_Average_CultureValues =  this.GetAverageHasValue(sourceRows, "Employee", "", "CultureValues");
            decimal vUnknown_Employee_Average_CareerOpportunities = this.GetAverageHasValue(sourceRows, "Employee", "", "CareerOpportunities");
            decimal vUnknown_Employee_Average_CompBenefits = this.GetAverageHasValue(sourceRows, "Employee", "", "CompBenefits");
            decimal vUnknown_Employee_Average_SeniorManagement = this.GetAverageHasValue(sourceRows, "Employee", "", "SeniorManagement");
            decimal vUnknown_Employee_CountNumber_Recommends =this.GetCountValue(sourceRows, "Employee", "", "Recommends", "Recommends");
            decimal vUnknown_Employee_CountNumber_Positive_Outlook = this.GetCountValue(sourceRows, "Employee", "", "Outlook", "Positive Outlook");
            decimal vUnknown_Employee_CountNumber_Negative_Outlook =  this.GetCountValue(sourceRows, "Employee", "", "Outlook", "Negative Outlook");
            decimal vUnknown_Employee_CountNumber_Neutral_Outlook =this.GetCountValue(sourceRows, "Employee", "", "Outlook", "Neutral Outlook");
            decimal vUnknown_Employee_CountNumber_NULL_Outlook = this.GetCountValueNullValue(sourceRows, "Employee", "", "Outlook");
            decimal vUnknown_Employee_CountNumber_No_opinion_of_CEO = this.GetCountValue(sourceRows, "Employee", "", "OptionOfCEO", "No opinion of CEO");
            decimal vUnknown_Employee_CountNumber_Approves_of_CEO = this.GetCountValue(sourceRows, "Employee", "", "OptionOfCEO", "Approves of CEO");
            decimal vUnknown_Employee_CountNumber_Disapproves_of_CEO = this.GetCountValue(sourceRows, "Employee", "", "OptionOfCEO", "Disapproves of CEO");
            decimal vUnknown_Employee_CountNumber_NULL_value_approves_CEO = this.GetCountValueNullValue(sourceRows, "Employee", "", "OptionOfCEO");

            //resultRow.Add("Key", key);
            //resultRow.Add("City", vCity); 
            //resultRow.Add("Year", vYear);
            resultRow.Add("Count_position_directors_managers", vCount_position_directors_managers);
            resultRow.Add("Count_position_others", vCount_position_others);
            resultRow.Add("Count_FullTime Job", vCount_FullTime_Job);
            resultRow.Add("Count_PartTime Job", vCount_PartTime_Job);
            resultRow.Add("Count_Ohters Job", vCount_Ohters_Job);
            resultRow.Add("Average_Words_Pros", vAverage_Words_Pros);
            resultRow.Add("Average_Words_Cons", vAverage_Words_Cons);
            resultRow.Add("Total_CountNumber_Rating", vTotal_CountNumber_Rating);
            resultRow.Add("Total_Average_Rating", vTotal_Average_Rating);
            resultRow.Add("Total_Average_WorkLifeBalance", vTotal_Average_WorkLifeBalance);
            resultRow.Add("Total_Average_CultureValues", vTotal_Average_CultureValues);
            resultRow.Add("Total_Average_CareerOpportunities", vTotal_Average_CareerOpportunities);
            resultRow.Add("Total_Average_CompBenefits", vTotal_Average_CompBenefits);
            resultRow.Add("Total_Average_SeniorManagement", vTotal_Average_SeniorManagement);
            resultRow.Add("Total_CountNumber_Recommends", vTotal_CountNumber_Recommends);
            resultRow.Add("Total_CountNumber_Positive Outlook", vTotal_CountNumber_Positive_Outlook);
            resultRow.Add("Total_CountNumber_Negative Outlook", vTotal_CountNumber_Negative_Outlook);
            resultRow.Add("Total_CountNumber_Neutral Outlook", vTotal_CountNumber_Neutral_Outlook);
            resultRow.Add("Total_CountNumber_NULL Outlook", vTotal_CountNumber_NULL_Outlook);
            resultRow.Add("Total_CountNumber_No opinion of CEO", vTotal_CountNumber_No_opinion_of_CEO);
            resultRow.Add("Total_CountNumber_Approves of CEO", vTotal_CountNumber_Approves_of_CEO);
            resultRow.Add("Total_CountNumber_Disapproves of CEO", vTotal_CountNumber_Disapproves_of_CEO);
            resultRow.Add("Total_CountNumber_NULL value approves_CEO", vTotal_CountNumber_NULL_value_approves_CEO);
            resultRow.Add("Former Employee_CountNumber_Rating", vFormer_Employee_CountNumber_Rating);
            resultRow.Add("Former Employee_Average_Rating", vFormer_Employee_Average_Rating);
            resultRow.Add("Former Employee_Average_WorkLifeBalance", vFormer_Employee_Average_WorkLifeBalance);
            resultRow.Add("Former Employee_Average_CultureValues", vFormer_Employee_Average_CultureValues);
            resultRow.Add("Former Employee_Average_CareerOpportunities", vFormer_Employee_Average_CareerOpportunities);
            resultRow.Add("Former Employee_Average_CompBenefits", vFormer_Employee_Average_CompBenefits);
            resultRow.Add("Former Employee_Average_SeniorManagement", vFormer_Employee_Average_SeniorManagement);
            resultRow.Add("Former Employee_CountNumber_Recommends", vFormer_Employee_CountNumber_Recommends);
            resultRow.Add("Former Employee_CountNumber_Positive Outlook", vFormer_Employee_CountNumber_Positive_Outlook);
            resultRow.Add("Former Employee_CountNumber_Negative Outlook", vFormer_Employee_CountNumber_Negative_Outlook);
            resultRow.Add("Former Employee_CountNumber_Neutral Outlook", vFormer_Employee_CountNumber_Neutral_Outlook);
            resultRow.Add("Former Employee_CountNumber_NULL Outlook", vFormer_Employee_CountNumber_NULL_Outlook);
            resultRow.Add("Former Employee_CountNumber_No opinion of CEO", vFormer_Employee_CountNumber_No_opinion_of_CEO);
            resultRow.Add("Former Employee_CountNumber_Approves of CEO", vFormer_Employee_CountNumber_Approves_of_CEO);
            resultRow.Add("Former Employee_CountNumber_Disapproves of CEO", vFormer_Employee_CountNumber_Disapproves_of_CEO);
            resultRow.Add("Former Employee_CountNumber_NULL value approves_CEO", vFormer_Employee_CountNumber_NULL_value_approves_CEO);
            resultRow.Add("Current Employee_CountNumber_Rating", vCurrent_Employee_CountNumber_Rating);
            resultRow.Add("Current Employee_Average_Rating", vCurrent_Employee_Average_Rating);
            resultRow.Add("Current Employee_Average_WorkLifeBalance", vCurrent_Employee_Average_WorkLifeBalance);
            resultRow.Add("Current Employee_Average_CultureValues", vCurrent_Employee_Average_CultureValues);
            resultRow.Add("Current Employee_Average_CareerOpportunities", vCurrent_Employee_Average_CareerOpportunities);
            resultRow.Add("Current Employee_Average_CompBenefits", vCurrent_Employee_Average_CompBenefits);
            resultRow.Add("Current Employee_Average_SeniorManagement", vCurrent_Employee_Average_SeniorManagement);
            resultRow.Add("Current Employee_CountNumber_Recommends", vCurrent_Employee_CountNumber_Recommends);
            resultRow.Add("Current Employee_CountNumber_Positive Outlook", vCurrent_Employee_CountNumber_Positive_Outlook);
            resultRow.Add("Current Employee_CountNumber_Negative Outlook", vCurrent_Employee_CountNumber_Negative_Outlook);
            resultRow.Add("Current Employee_CountNumber_Neutral Outlook", vCurrent_Employee_CountNumber_Neutral_Outlook);
            resultRow.Add("Current Employee_CountNumber_NULL Outlook", vCurrent_Employee_CountNumber_NULL_Outlook);
            resultRow.Add("Current Employee_CountNumber_No opinion of CEO", vCurrent_Employee_CountNumber_No_opinion_of_CEO);
            resultRow.Add("Current Employee_CountNumber_Approves of CEO", vCurrent_Employee_CountNumber_Approves_of_CEO);
            resultRow.Add("Current Employee_CountNumber_Disapproves of CEO", vCurrent_Employee_CountNumber_Disapproves_of_CEO);
            resultRow.Add("Current Employee_CountNumber_NULL value approves_CEO", vCurrent_Employee_CountNumber_NULL_value_approves_CEO);
            resultRow.Add("Unknown Employee_CountNumber_Rating", vUnknown_Employee_CountNumber_Rating);
            resultRow.Add("Unknown Employee_Average_Rating", vUnknown_Employee_Average_Rating);
            resultRow.Add("Unknown Employee_Average_WorkLifeBalance", vUnknown_Employee_Average_WorkLifeBalance);
            resultRow.Add("Unknown Employee_Average_CultureValues", vUnknown_Employee_Average_CultureValues);
            resultRow.Add("Unknown Employee_Average_CareerOpportunities", vUnknown_Employee_Average_CareerOpportunities);
            resultRow.Add("Unknown Employee_Average_CompBenefits", vUnknown_Employee_Average_CompBenefits);
            resultRow.Add("Unknown Employee_Average_SeniorManagement", vUnknown_Employee_Average_SeniorManagement);
            resultRow.Add("Unknown Employee_CountNumber_Recommends", vUnknown_Employee_CountNumber_Recommends);
            resultRow.Add("Unknown Employee_CountNumber_Positive Outlook", vUnknown_Employee_CountNumber_Positive_Outlook);
            resultRow.Add("Unknown Employee_CountNumber_Negative Outlook", vUnknown_Employee_CountNumber_Negative_Outlook);
            resultRow.Add("Unknown Employee_CountNumber_Neutral Outlook", vUnknown_Employee_CountNumber_Neutral_Outlook);
            resultRow.Add("Unknown Employee_CountNumber_NULL Outlook", vUnknown_Employee_CountNumber_NULL_Outlook);
            resultRow.Add("Unknown Employee_CountNumber_No opinion of CEO", vUnknown_Employee_CountNumber_No_opinion_of_CEO);
            resultRow.Add("Unknown Employee_CountNumber_Approves of CEO", vUnknown_Employee_CountNumber_Approves_of_CEO);
            resultRow.Add("Unknown Employee_CountNumber_Disapproves of CEO", vUnknown_Employee_CountNumber_Disapproves_of_CEO);
            resultRow.Add("Unknown Employee_CountNumber_NULL value approves_CEO", vUnknown_Employee_CountNumber_NULL_value_approves_CEO);

            return resultRow;
        }

        private decimal GetCountValue(List<Dictionary<string, string>> sourceRows, string columnName, string matchValue)
        {
            decimal count = 0;
            foreach (Dictionary<string, string> sourceRow in sourceRows)
            {
                string value = sourceRow[columnName];
                if (value == matchValue)
                {
                    count++;
                }
            }
            return count;
        }


        private decimal GetCountValueHasValue(List<Dictionary<string, string>> sourceRows, string columnName)
        {
            decimal count = 0;
            foreach (Dictionary<string, string> sourceRow in sourceRows)
            {
                string value = sourceRow[columnName].Trim();
                if (value.Length > 0)
                {
                    count++;
                }
            }
            return count;
        }


        private decimal GetCountValueNullValue(List<Dictionary<string, string>> sourceRows, string columnName)
        {
            decimal count = 0;
            foreach (Dictionary<string, string> sourceRow in sourceRows)
            {
                string value = sourceRow[columnName].Trim();
                if (value.Length == 0)
                {
                    count++;
                }
            }
            return count;
        }

        private decimal GetAverageValue(List<Dictionary<string, string>> sourceRows, string columnName)
        {
            decimal sum = 0;
            decimal count = 0;
            foreach (Dictionary<string, string> sourceRow in sourceRows)
            {
                decimal value = decimal.Parse(sourceRow[columnName]);
                sum = sum + value;
                count++;
            }
            return sourceRows.Count == 0 ? 0 : (sum / count);
        }

        private decimal GetAverageHasValue(List<Dictionary<string, string>> sourceRows, string columnName)
        {
            decimal sum = 0;
            decimal count = 0;
            foreach (Dictionary<string, string> sourceRow in sourceRows)
            {
                string valueStr = sourceRow[columnName].Trim();
                if (valueStr.Length > 0)
                {
                    decimal value = decimal.Parse(valueStr);
                    sum = sum + value;
                    count++;
                }
            }
            return count == 0 ? 0 : (sum / count);
        }




        private decimal GetCountValue(List<Dictionary<string, string>> sourceRows, string filterColumnName, string filterMatchValue, string columnName, string matchValue)
        {
            decimal count = 0;
            foreach (Dictionary<string, string> sourceRow in sourceRows)
            {
                string filterValue = sourceRow[filterColumnName];
                string value = sourceRow[columnName];
                if (filterValue == filterMatchValue && value == matchValue)
                {
                    count++;
                }
            }
            return count;
        }


        private decimal GetCountValueHasValue(List<Dictionary<string, string>> sourceRows, string filterColumnName, string filterMatchValue, string columnName)
        {
            decimal count = 0;
            foreach (Dictionary<string, string> sourceRow in sourceRows)
            {
                string filterValue = sourceRow[filterColumnName];
                string value = sourceRow[columnName].Trim();
                if (filterValue == filterMatchValue && value.Length > 0)
                {
                    count++;
                }
            }
            return count;
        }


        private decimal GetCountValueNullValue(List<Dictionary<string, string>> sourceRows, string filterColumnName, string filterMatchValue, string columnName)
        {
            decimal count = 0;
            foreach (Dictionary<string, string> sourceRow in sourceRows)
            {
                string filterValue = sourceRow[filterColumnName];
                string value = sourceRow[columnName].Trim();
                if (filterValue == filterMatchValue && value.Length == 0)
                {
                    count++;
                }
            }
            return count;
        }

        private decimal GetAverageValue(List<Dictionary<string, string>> sourceRows, string filterColumnName, string filterMatchValue, string columnName)
        {
            decimal sum = 0;
            decimal count = 0;
            foreach (Dictionary<string, string> sourceRow in sourceRows)
            {
                string filterValue = sourceRow[filterColumnName];
                if (filterValue == filterMatchValue)
                {
                    decimal value = decimal.Parse(sourceRow[columnName]);
                    sum = sum + value;
                    count++;
                }
            }
            return sourceRows.Count == 0 ? 0 : (sum / count);
        }

        private decimal GetAverageHasValue(List<Dictionary<string, string>> sourceRows, string filterColumnName, string filterMatchValue, string columnName)
        {
            decimal sum = 0;
            decimal count = 0;
            foreach (Dictionary<string, string> sourceRow in sourceRows)
            {
                string filterValue = sourceRow[filterColumnName];
                string valueStr = sourceRow[columnName].Trim();
                if (filterValue == filterMatchValue && valueStr.Length > 0)
                {
                    decimal value = decimal.Parse(valueStr);
                    sum = sum + value;
                    count++;
                }
            }
            return count == 0 ? 0 : (sum / count);
        }

        private void ProcessWRDSCompanyCityYearList(string sourceFilePath, string destFilePath)
        {
            XSSFWorkbook wb = new XSSFWorkbook(sourceFilePath);
            ExcelReader listSummaryEr = new ExcelReader(wb, "ListSummary");
            int listSummaryRowCount = listSummaryEr.GetRowCount();

            ExcelReader reviewEr = new ExcelReader(wb, "List-detai_review_new");
            int reviewRowCount = reviewEr.GetRowCount();

            ExcelReader wrdsEr = new ExcelReader(wb, "WRDS");
            int wrdsRowCount = wrdsEr.GetRowCount();

            Dictionary<string, Dictionary<string, string>> wrdsCompany2CompanyInfo = new Dictionary<string, Dictionary<string, string>>();
            Dictionary<string, Dictionary<string,string>> glassDoorCompany2CompanInfo = new Dictionary<string,  Dictionary<string,string>>();
            for (int i = 0; i < listSummaryRowCount; i++)
            {
                Dictionary<string, string> row = listSummaryEr.GetFieldValues(i);
                string wrdsCompanyKey = row["Auditor fkey"];
                string wrdsCompany = row["Auditor name"];
                string sourceCompany = row["Source Company Name"];
                string searchCompany = row["Company Name"];
                string glassDoorCompany = row["Page_Company_Name"];
                if (wrdsCompany.Length > 0)
                {
                    Dictionary<string, string> companInfo = new Dictionary<string, string>();
                    companInfo.Add("GlassDoor Company Name", glassDoorCompany);
                    companInfo.Add("Auditor fkey", wrdsCompanyKey);
                    wrdsCompany2CompanyInfo.Add(wrdsCompany, companInfo);
                }
                {
                    Dictionary<string, string> companInfo = new Dictionary<string, string>();
                    companInfo.Add("Source company name", sourceCompany);
                    companInfo.Add("Search company name", searchCompany);
                    companInfo.Add("GlassDoor company name", glassDoorCompany);
                    companInfo.Add("WRDS company name", wrdsCompany);
                    glassDoorCompany2CompanInfo.Add(glassDoorCompany, companInfo);
                }
            }

            Dictionary<string, List<Dictionary<string, string>>> wrdsCities = new Dictionary<string, List<Dictionary<string, string>>>();
            Dictionary<string, List<string>> wrdsCompanyCityDic = new Dictionary<string, List<string>>();
            List<string> wrdsCompanyNames = new List<string>();
            for (int i = 0; i < wrdsRowCount; i++)
            {
                Dictionary<string, string> row = wrdsEr.GetFieldValues(i);
                string wrdsCompanyName = row["Auditor name"];
                string city = row["city_state"];

                if (!wrdsCompanyCityDic.ContainsKey(wrdsCompanyName))
                {
                    wrdsCompanyCityDic.Add(wrdsCompanyName, new List<string>());
                }
                List<string> cities = wrdsCompanyCityDic[wrdsCompanyName];
                if (!cities.Contains(city))
                {
                    cities.Add(city);
                }

                if (wrdsCompany2CompanyInfo.ContainsKey(wrdsCompanyName))
                {
                    if (!wrdsCities.ContainsKey(wrdsCompanyName))
                    {
                        wrdsCompanyNames.Add(wrdsCompanyName);
                        wrdsCities.Add(wrdsCompanyName, new List<Dictionary<string, string>>());
                    }
                    List<Dictionary<string, string>> rows = wrdsCities[wrdsCompanyName];
                    rows.Add(row);
                }
            }

            Dictionary<string, List<string>> glassDoorCompanyCityDic = new Dictionary<string, List<string>>();
            Dictionary<string, Dictionary<string, object>> glassDoorCompanyCityYears = new Dictionary<string, Dictionary<string, object>>();
            List<string> glassDoorCompanyNames = new List<string>();
            for (int i = 0; i < reviewRowCount; i++)
            {
                Dictionary<string, string> row = reviewEr.GetFieldValues(i);
                string company = row["Page_Company_Name"];
                string city = row["Location"];
                string year = row["year"];

                if (!glassDoorCompanyNames.Contains(company))
                {
                    glassDoorCompanyNames.Add(company);
                }

                if (!glassDoorCompanyCityDic.ContainsKey(company))
                { 
                    glassDoorCompanyCityDic.Add(company, new List<string>());
                }

                List<string> cities = glassDoorCompanyCityDic[company];
                if (city.Length != 0 && !cities.Contains(city))
                {
                    cities.Add(city);
                }

                string key = company + "_" + city + "_" + year;
                if (!glassDoorCompanyCityYears.ContainsKey(key))
                {
                    Dictionary<string, object> newValue = new Dictionary<string, object>();
                    newValue.Add("Company", company);
                    newValue.Add("City", city);
                    newValue.Add("Year", year);
                    newValue.Add("Rows", new List<Dictionary<string, string>>());
                    glassDoorCompanyCityYears.Add(key, newValue);
                }
                Dictionary<string, object> value = glassDoorCompanyCityYears[key];
                List<Dictionary<string, string>> rows = (List<Dictionary<string, string>>)value["Rows"];
                rows.Add(row);
            }            

            ExcelWriter resultEW = this.GetExcelWriter(destFilePath);

            for (int i = 0; i < wrdsCompanyNames.Count; i++)
            {
                string wrdsCompanyName = wrdsCompanyNames[i];
                Dictionary<string, string> wrdsCompanyInfo = wrdsCompany2CompanyInfo[wrdsCompanyName];
                string glassDoorCompanyName = wrdsCompanyInfo["GlassDoor Company Name"];
                string wrdsAuditor_fkey = wrdsCompanyInfo["Auditor fkey"];
                List<Dictionary<string, string>> wrdsRows = wrdsCities[wrdsCompanyName];
                for (decimal year = 2008; year <= 2018; year++)
                {
                    for (int j = 0; j < wrdsRows.Count; j++)
                    {
                        Dictionary<string, string> wrdsRow = wrdsRows[j];
                        string vAuditor_fkey = wrdsRow["Auditor fkey"];
                        string vAuditor_name = wrdsRow["Auditor name"];
                        string vAuditor_city = wrdsRow["Auditor city"];
                        string vAuditor_state_name = wrdsRow["Auditor state name"];
                        string vAuditor_state = wrdsRow["Auditor state"];
                        string vAuditor_country = wrdsRow["Auditor country"];
                        string vAuditor_region = wrdsRow["Auditor region"];
                        string vcity_state = wrdsRow["city_state"];

                        string key = glassDoorCompanyName + "_" + vcity_state + "_" + year.ToString();

                        Dictionary<string, object> statisticRow = null;
                        if (glassDoorCompanyCityYears.ContainsKey(key))
                        {
                            Dictionary<string, object> reviewCity = glassDoorCompanyCityYears[key];
                            List<Dictionary<string, string>> reviewRows = (List<Dictionary<string, string>>)reviewCity["Rows"];
                            statisticRow = this.ProcessWRDSCompanyCityYearRow(key, glassDoorCompanyName, vcity_state, year.ToString(), reviewRows);
                        }
                        else
                        {
                            statisticRow = new Dictionary<string, object>();
                        }
                        statisticRow["City_State"] = vcity_state;

                        statisticRow["Auditor fkey"] = vAuditor_fkey;
                        statisticRow["Auditor name"] = vAuditor_name;
                        statisticRow["Auditor city"] = vAuditor_city;
                        statisticRow["Auditor state name"] = vAuditor_state_name;
                        statisticRow["Auditor state"] = vAuditor_state;
                        statisticRow["Auditor country"] = vAuditor_country;
                        statisticRow["Auditor region"] = vAuditor_region;
                        statisticRow["Year"] = year;

                        Dictionary<string, string> companInfo = glassDoorCompany2CompanInfo[glassDoorCompanyName];
                        statisticRow["Source company name"] = companInfo["Source company name"];
                        statisticRow["Search company name"] = companInfo["Search company name"];
                        statisticRow["GlassDoor company name"] = companInfo["GlassDoor company name"];

                        statisticRow["CompanyInWRDS"] = "True";
                        statisticRow["CompanyCityInWRDS"] = "True";

                        resultEW.AddRow(statisticRow);
                    }

                    List<string> glassDoorCompanyCities = glassDoorCompanyCityDic[glassDoorCompanyName];
                    List<string> wrdsCompanyCities = wrdsCompanyCityDic[wrdsCompanyName];
                    {
                        //在wrds里没有出现的城市
                        for (int j = 0; j < glassDoorCompanyCities.Count; j++)
                        {
                            string glassDoorCity = glassDoorCompanyCities[j];
                            if (!wrdsCompanyCities.Contains(glassDoorCity))
                            {
                                string key = glassDoorCompanyName + "_" + glassDoorCity + "_" + year.ToString();

                                Dictionary<string, object> statisticRow = null;
                                if (glassDoorCompanyCityYears.ContainsKey(key))
                                {
                                    Dictionary<string, object> reviewCity = glassDoorCompanyCityYears[key];
                                    List<Dictionary<string, string>> reviewRows = (List<Dictionary<string, string>>)reviewCity["Rows"];
                                    statisticRow = this.ProcessWRDSCompanyCityYearRow(key, glassDoorCompanyName, glassDoorCity, year.ToString(), reviewRows);
                                }
                                else
                                {
                                    statisticRow = new Dictionary<string, object>();
                                }
                                statisticRow["City_State"] = glassDoorCity;

                                statisticRow["Auditor fkey"] = wrdsAuditor_fkey;
                                statisticRow["Auditor name"] = wrdsCompanyName;
                                statisticRow["Year"] = year;

                                Dictionary<string, string> companInfo = glassDoorCompany2CompanInfo[glassDoorCompanyName];
                                statisticRow["Source company name"] = companInfo["Source company name"];
                                statisticRow["Search company name"] = companInfo["Search company name"];
                                statisticRow["GlassDoor company name"] = companInfo["GlassDoor company name"];

                                statisticRow["CompanyInWRDS"] = "True";
                                statisticRow["CompanyCityInWRDS"] = "False";

                                resultEW.AddRow(statisticRow);
                            }
                        }
                        {
                            //城市名为空
                            string glassDoorCity = "";
                            string key = glassDoorCompanyName + "_" + glassDoorCity + "_" + year.ToString();

                            Dictionary<string, object> statisticRow = null;
                            if (glassDoorCompanyCityYears.ContainsKey(key))
                            {
                                Dictionary<string, object> reviewCity = glassDoorCompanyCityYears[key];
                                List<Dictionary<string, string>> reviewRows = (List<Dictionary<string, string>>)reviewCity["Rows"];
                                statisticRow = this.ProcessWRDSCompanyCityYearRow(key, glassDoorCompanyName, glassDoorCity, year.ToString(), reviewRows);
                            }
                            else
                            {
                                statisticRow = new Dictionary<string, object>();
                            }
                            statisticRow["City_State"] = glassDoorCity;

                            statisticRow["Auditor fkey"] = wrdsAuditor_fkey;
                            statisticRow["Auditor name"] = wrdsCompanyName;
                            statisticRow["Year"] = year;

                            Dictionary<string, string> companInfo = glassDoorCompany2CompanInfo[glassDoorCompanyName];
                            statisticRow["Source company name"] = companInfo["Source company name"];
                            statisticRow["Search company name"] = companInfo["Search company name"];
                            statisticRow["GlassDoor company name"] = companInfo["GlassDoor company name"];

                            statisticRow["CompanyInWRDS"] = "True";
                            statisticRow["CompanyCityInWRDS"] = "False";

                            resultEW.AddRow(statisticRow); 
                        }
                    }
                }
            }

            for (int i = 0; i < glassDoorCompanyNames.Count; i++)
            {
                string glassDoorCompanyName = glassDoorCompanyNames[i];
                Dictionary<string, string> glassDoorCompanyInfo = glassDoorCompany2CompanInfo[glassDoorCompanyName];
                string wrdsCompanyName = glassDoorCompanyInfo["WRDS company name"];
                if (wrdsCompanyName.Length == 0)
                {
                    List<string> cities = glassDoorCompanyCityDic[glassDoorCompanyName];
                    for (decimal year = 2008; year <= 2018; year++)
                    {
                        {
                            //存在城市名
                            for (int j = 0; j < cities.Count; j++)
                            {
                                string city = cities[j];

                                string key = glassDoorCompanyName + "_" + city + "_" + year.ToString();

                                Dictionary<string, object> statisticRow = null;
                                if (glassDoorCompanyCityYears.ContainsKey(key))
                                {
                                    Dictionary<string, object> reviewCity = glassDoorCompanyCityYears[key];
                                    List<Dictionary<string, string>> reviewRows = (List<Dictionary<string, string>>)reviewCity["Rows"];
                                    statisticRow = this.ProcessWRDSCompanyCityYearRow(key, glassDoorCompanyName, city, year.ToString(), reviewRows);
                                }
                                else
                                {
                                    statisticRow = new Dictionary<string, object>();
                                }
                                statisticRow["City_State"] = city;
                                statisticRow["Year"] = year;

                                Dictionary<string, string> companInfo = glassDoorCompany2CompanInfo[glassDoorCompanyName];
                                statisticRow["Source company name"] = companInfo["Source company name"];
                                statisticRow["Search company name"] = companInfo["Search company name"];
                                statisticRow["GlassDoor company name"] = companInfo["GlassDoor company name"];

                                statisticRow["CompanyInWRDS"] = "False";
                                statisticRow["CompanyCityInWRDS"] = "False";

                                resultEW.AddRow(statisticRow);
                            }
                        }
                        {
                            //城市名为空
                            string city = "";

                            string key = glassDoorCompanyName + "_" + city + "_" + year.ToString();

                            Dictionary<string, object> statisticRow = null;
                            if (glassDoorCompanyCityYears.ContainsKey(key))
                            {
                                Dictionary<string, object> reviewCity = glassDoorCompanyCityYears[key];
                                List<Dictionary<string, string>> reviewRows = (List<Dictionary<string, string>>)reviewCity["Rows"];
                                statisticRow = this.ProcessWRDSCompanyCityYearRow(key, glassDoorCompanyName, city, year.ToString(), reviewRows);
                            }
                            else
                            {
                                statisticRow = new Dictionary<string, object>();
                            }
                            statisticRow["City_State"] = city;
                            statisticRow["Year"] = year;

                            Dictionary<string, string> companInfo = glassDoorCompany2CompanInfo[glassDoorCompanyName];
                            statisticRow["Source company name"] = companInfo["Source company name"];
                            statisticRow["Search company name"] = companInfo["Search company name"];
                            statisticRow["GlassDoor company name"] = companInfo["GlassDoor company name"];

                            statisticRow["CompanyInWRDS"] = "False";
                            statisticRow["CompanyCityInWRDS"] = "False";

                            resultEW.AddRow(statisticRow);
                        }
                    }
                }

            }

            resultEW.SaveToDisk();
        } 
    }
}