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

namespace NetDataAccess.Extended.Jianzhu.JinanLouPan
{
    public class GetHousePages : ExternalRunWebPage
    { 
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            try
            {
                return this.GetHouseListInfos(listSheet) ;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        private CsvWriter CreateResultCsvWriter()
        {
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>(); 
            resultColumnDic.Add("projectId", 0);
            resultColumnDic.Add("项目名称", 1);
            resultColumnDic.Add("buildingId", 2);
            resultColumnDic.Add("楼名称", 3);
            resultColumnDic.Add("是否住宅房屋", 4);
            resultColumnDic.Add("单元号", 5);
            resultColumnDic.Add("顺序号", 6);
            resultColumnDic.Add("楼层", 7);
            resultColumnDic.Add("houseId", 8);
            resultColumnDic.Add("houseName", 9);
            resultColumnDic.Add("房屋面积", 10);
            resultColumnDic.Add("套内面积", 11);
            resultColumnDic.Add("公摊面积", 12);
            resultColumnDic.Add("房屋用途", 13);
            resultColumnDic.Add("销售状态编码", 14);
            resultColumnDic.Add("销售状态", 15); 
            string resultFilePath = Path.Combine(exportDir, "济南楼盘_房间信息.csv");
            Dictionary<string, string> resultColumnFormat = new Dictionary<string, string>();
            CsvWriter resultEW = new CsvWriter(resultFilePath, resultColumnDic);
            return resultEW;
        }

        private bool GetHouseListInfos(IListSheet listSheet)
        {

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, string> houseBuildingDic = new Dictionary<string, string>();
             
            CsvWriter resultEW = this.CreateResultCsvWriter();

            for (int i = 0; i < listSheet.RowCount; i++)
            { 
                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string pageUrl = listSheet.PageUrlList[i];
                    string localFilePath = this.RunPage.GetFilePath(pageUrl, pageSourceDir);
                    string fileText = FileHelper.GetTextFromFile(localFilePath);
                    JObject rootJo = JObject.Parse(fileText);
                    string houseStatusNo = rootJo["housestatus"].ToString();
                    string fid = rootJo["fid"].ToString();
                    string unitArea = rootJo["unitarea"].ToString();
                    string apportioArea = rootJo["apportioarea"].ToString();
                    string usedTypeNo = rootJo["usedtypeno"].ToString();
                    string houseArea = rootJo["housearea"].ToString();
                    string houseStatus = "";
                    switch (houseStatusNo)
                    {
                        case "15701": 
                            houseStatus = "可售"; 
                            break;
                        case "15702":
                            houseStatus = "已预订";
                            break;
                        case "15703":
                            houseStatus = "已备案";
                            break;
                        case "15704": 
                            houseStatus = "已签约";
                            break;
                        case "15705":
                            houseStatus = "可租"; 
                            break;
                        case "15707": 
                            houseStatus = "不可租售";
                            break;
                        case "15709":
                            houseStatus = "已预订";
                            break;
                        case "15710": 
                            houseStatus = "查封"; 
                            break;
                        case "15711":
                            houseStatus = "冻结"; 
                            break;
                        default:
                            houseStatus = "可售"; 
                            break;
                    }


                    Dictionary<string, string> f2vs = new Dictionary<string, string>();
                    f2vs.Add("projectId", row["projectId"]);
                    f2vs.Add("项目名称", row["项目名称"]);
                    f2vs.Add("buildingId", row["buildingId"]);
                    f2vs.Add("楼名称", row["楼名称"]);
                    f2vs.Add("是否住宅房屋", row["是否住宅房屋"]);
                    f2vs.Add("单元号", row["单元号"]);
                    f2vs.Add("顺序号", row["顺序号"]);
                    f2vs.Add("楼层", row["楼层"]);
                    f2vs.Add("houseId", row["houseId"]);
                    f2vs.Add("houseName", row["houseName"]);
                    f2vs.Add("房屋面积", houseArea);
                    f2vs.Add("套内面积", unitArea);
                    f2vs.Add("公摊面积", apportioArea);
                    f2vs.Add("房屋用途", usedTypeNo);
                    f2vs.Add("销售状态编码", houseStatusNo);
                    f2vs.Add("销售状态", houseStatus); 
                    resultEW.AddRow(f2vs);
                }
            }

            resultEW.SaveToDisk();

            return true;
        }
         
    }
}