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
    public class GetListPageUrls : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            this.GenerateListPageUrls();
            return true;
        }

        private ExcelWriter GetExcelWriter(string destFilePath)
        { 

            Dictionary<string, int> columnDic = CommonUtil.InitStringIndexDic(new string[]{
                    "detailPageUrl",
                    "detailPageName", 
                    "cookie",
                    "grabStatus", 
                    "giveUpGrab", 
                    "Company_Name"});

            ExcelWriter ew = new ExcelWriter(destFilePath, "List", columnDic);
            return ew;
        }

        private void GenerateListPageUrls()
        {
            string[] parameters = this.Parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            string sourceFilePath = parameters[0];
            string destFilePath = parameters[1];

            ExcelReader er = new ExcelReader(sourceFilePath);
            ExcelWriter ew = this.GetExcelWriter(destFilePath);

            Dictionary<string, string> companyDic = new Dictionary<string, string>();

            int rowCount = er.GetRowCount();
            for (int i = 0; i < rowCount; i++)
            {
                Dictionary<string, string> sourceRow = er.GetFieldValues(i);

                string companyName = sourceRow["Company Name"];
                if (!companyDic.ContainsKey(companyName))
                {
                    companyDic.Add(companyName, null);

                    string encodeCompanyName = CommonUtil.UrlEncode(companyName);
                    string pageUrl = "https://www.glassdoor.com/Reviews/company-reviews.htm?suggestCount=0&suggestChosen=false&clickSource=searchBtn&typedKeyword=" + encodeCompanyName + "&sc.keyword=" + encodeCompanyName + "&locT=&locId=&jobType=";
                    Dictionary<string, string> destRow = new Dictionary<string, string>();
                    destRow.Add("detailPageUrl", pageUrl);
                    destRow.Add("detailPageName", pageUrl);
                    destRow.Add("cookie", "ARPNTS=1952819392.64288.0000; ARPNTS_AB=115; gdId=94517b85-9d89-47c1-a5ab-2a04b242c067; trs=direct:direct:direct:2018-07-15+23%3A50%3A22.919:undefined:undefined; _ga=GA1.2.216399378.1531723803; __qca=P0-1262345758-1531723804448; G_ENABLED_IDPS=google; __gads=ID=62251b7c5d596d61:T=1531723836:S=ALNI_MZk81H-OcTT9PjdVFK8PYIrVGTx1A; __gdpopuc=1; cto_lwid=8e5c6f44-854b-492e-be0f-09a9dc915819; rm=bGl4aW4xNTUzQGdtYWlsLmNvbToxNTYzMjkyNzgzNzgxOjVhMDQ1MWI1NjBiYjYzYzE3NjM3YmEzOThjNTJlM2Ix; uc=8F0D0CFA50133D96DAB3D34ABA1B873399807652C6C76982808553CADAB58BBB131EFE7DE1E6A4B95851EB3294212EB393007ED539985D9CDE873DE04D4FC71FEE18FB9F0BDE4138B3E34D8411CDEA90F25EDE93274F0D5D5FDED9B003FBA6F43CA9014AC0BB0289EB0204D279873038C3CF7E94AE6F099E0174A86BB3453633759C8511C218159EA514952BE5A78210E84BCCC56AAAAD09; _mibhv=anon-1531735166141-5684441656_6890; JSESSIONID=E01A35E3A52310CD24E42EC5FF252052; _uac=00000164ca63c0a6a0bcb163ea7dc134; GSESSIONID=E01A35E3A52310CD24E42EC5FF252052; _gid=GA1.2.739342608.1532403870; ht=%7B%22quantcast%22%3A%5B%22D%22%5D%7D; JSESSIONID_KYWI_APP=B31D8DA6C274B6196C84875AE7D7942A; JSESSIONID_JX_APP=8E738CEAE7DF1A613C3E7B6006442DE4; cass=1; AWSALB=4plUYq9nqfzCEW/AJ4UDiC11DqFrHS0JteBY5hN5Ok2HoX9iLI04hye/Bpq8j7Syv8PnKkRAsMcWCNGXkxlGMnVvbn1nPp99yMD5TcSM4g+ORjkL9rbNGIoiSAN4qYGv/Ir11PRBDXxXJIp8E0TRnpuNo3fcuCImeBiC/rzpGYMOeTyaTw32g+C3rlEvLCmoohAKaUTrzpDwu1OXM6sBFG9S5jfl0NtU/cmqv5muVCjMBcqr8FoqD9WjlkUkNe0=");
                    destRow.Add("Company_Name", companyName);
                    ew.AddRow(destRow);
                }
            }
            ew.SaveToDisk(); 
        }
    }
}