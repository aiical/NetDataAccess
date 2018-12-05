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

namespace NetDataAccess.Extended.GuPiao.LQF
{
    /// <summary>
    /// 匹配招股说明书
    /// </summary>
    public class MatchZGSMS : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            try
            {
                this.DoMatch(listSheet);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return true;
        }

        private void Sort(List<Dictionary<string, string>> announcementList)
        {
            DateTime maxTime = new DateTime(1900,1,1);
            List<Dictionary<string, string>> sortedList = new List<Dictionary<string, string>>();
            for (int i = 0; i < announcementList.Count; i++)
            {
                Dictionary<string,string> tempAnnouncement = announcementList[i];
                DateTime tempTime = DateTime.Parse(tempAnnouncement["发布日期"]);
                int insertPosition = sortedList.Count;
                for (int j = 0; j < sortedList.Count; j++)
                {
                    Dictionary<string, string> checkAnnouncement = sortedList[j];
                    DateTime checkTime = DateTime.Parse(checkAnnouncement["发布日期"]);
                    if (checkTime < tempTime)
                    {
                        insertPosition = j;
                    }
                }
                if (insertPosition >= sortedList.Count)
                {
                    sortedList.Add(tempAnnouncement);
                }
                else
                {
                    sortedList.Insert(insertPosition, tempAnnouncement);
                }
            }
        }


        private void DoMatch(IListSheet listSheet)
        {
            string[] parameters = this.Parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            string sourceFilePath = parameters[0];
            string destFilePath = parameters[1];

            ExcelReader stkListER = new ExcelReader(sourceFilePath, "stkList");
            int stkListRowCount = stkListER.GetRowCount();


            ExcelReader announcementListER = new ExcelReader(sourceFilePath, "announcementList");
            int announcementListRowCount = announcementListER.GetRowCount();

            Dictionary<string, List<Dictionary<string, string>>> codeToAnnouncementListDic = new Dictionary<string, List<Dictionary<string, string>>>();
            for (int i = 0; i < announcementListRowCount; i++)
            {
                Dictionary<string, string> announcementRow = announcementListER.GetFieldValues(i);
                string code = announcementRow["编码"];
                if (codeToAnnouncementListDic.ContainsKey(code))
                {
                    codeToAnnouncementListDic[code].Add(announcementRow);
                }
                else
                {
                    codeToAnnouncementListDic.Add(code, new List<Dictionary<string, string>>() { announcementRow });
                }
            }
            
            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("Stkcd", 0);
            resultColumnDic.Add("Stknme", 1);
            resultColumnDic.Add("Listexg", 2);
            resultColumnDic.Add("Estbdt", 3);
            resultColumnDic.Add("Ipodt", 4);
            resultColumnDic.Add("Listdt", 5);
            resultColumnDic.Add("year", 6);
            resultColumnDic.Add("Notes", 7);
            resultColumnDic.Add("简写", 8);
            resultColumnDic.Add("名称", 9);
            resultColumnDic.Add("编码", 10);
            resultColumnDic.Add("交易所", 11);
            resultColumnDic.Add("类型", 12);
            resultColumnDic.Add("标题", 13);
            resultColumnDic.Add("发布日期", 14);
            resultColumnDic.Add("url", 15);
            resultColumnDic.Add("单次招股的说明书个数", 16);
            Dictionary<string, string> columnFormats = new Dictionary<string, string>();
            columnFormats.Add("Listexg", "#0");
            columnFormats.Add("Estbdt", "yyyy/m/d");
            columnFormats.Add("Ipodt", "yyyy/m/d");
            columnFormats.Add("Listdt", "yyyy/m/d");
            columnFormats.Add("year", "#0");
            columnFormats.Add("发布日期", "yyyy/m/d");
            ExcelWriter resultEW = new ExcelWriter(destFilePath, "Matched", resultColumnDic, columnFormats);

            for (int i = 0; i < stkListRowCount; i++)
            {
                Dictionary<string, string> stkRow = stkListER.GetFieldValues(i);
                string stkcd = stkRow["Stkcd"];

                int matchedCount = 0;
                if (codeToAnnouncementListDic.ContainsKey(stkcd))
                {
                    DateTime listdt = DateTime.Parse(stkRow["Listdt"]);

                    List<Dictionary<string, string>> announcementRows = codeToAnnouncementListDic[stkcd];
                    List<Dictionary<string, string>> matchedAnnouncementRows = new List<Dictionary<string, string>>();
                    for (int j = 0; j < announcementRows.Count; j++)
                    {
                        Dictionary<string, string> announcementRow = announcementRows[j]; 
                        DateTime announcementDate = DateTime.Parse(announcementRow["发布日期"]);
                        //存在特殊情况，发布招股说明书前进行招股
                        if (listdt.AddDays(30) >= announcementDate && listdt.AddDays(-365) < announcementDate)
                        {
                            matchedCount++;
                            matchedAnnouncementRows.Add(announcementRow);
                        }
                    }
                    if (matchedAnnouncementRows.Count == 0)
                    {
                        Dictionary<string, object> resultRow = new Dictionary<string, object>();
                        resultRow.Add("Stkcd", stkRow["Stkcd"]);
                        resultRow.Add("Stknme", stkRow["Stknme"]);
                        resultRow.Add("Listexg", int.Parse(stkRow["Listexg"]));
                        resultRow.Add("Estbdt", DateTime.Parse(stkRow["Estbdt"]));
                        resultRow.Add("Ipodt", DateTime.Parse(stkRow["Ipodt"]));
                        resultRow.Add("Listdt", DateTime.Parse(stkRow["Listdt"]));
                        resultRow.Add("year", int.Parse(stkRow["year"]));
                        resultRow.Add("Notes", stkRow["Notes"]);
                        resultRow.Add("单次招股的说明书个数", matchedCount);
                        resultEW.AddRow(resultRow);
                    }
                    else
                    {
                        foreach (Dictionary<string, string> announcementRow in matchedAnnouncementRows)
                        {

                            Dictionary<string, object> resultRow = new Dictionary<string, object>();
                            resultRow.Add("Stkcd", stkRow["Stkcd"]);
                            resultRow.Add("Stknme", stkRow["Stknme"]);
                            resultRow.Add("Listexg", int.Parse(stkRow["Listexg"]));
                            resultRow.Add("Estbdt", DateTime.Parse(stkRow["Estbdt"]));
                            resultRow.Add("Ipodt", DateTime.Parse(stkRow["Ipodt"]));
                            resultRow.Add("Listdt", DateTime.Parse(stkRow["Listdt"]));
                            resultRow.Add("year", int.Parse(stkRow["year"]));
                            resultRow.Add("Notes", stkRow["Notes"]);
                            resultRow.Add("简写", announcementRow["简写"]);
                            resultRow.Add("名称", announcementRow["名称"]);
                            resultRow.Add("编码", announcementRow["编码"]);
                            resultRow.Add("交易所", announcementRow["交易所"]);
                            resultRow.Add("类型", announcementRow["类型"]);
                            resultRow.Add("标题", announcementRow["标题"]);
                            resultRow.Add("发布日期",  DateTime.Parse(announcementRow["发布日期"]));
                            resultRow.Add("url", announcementRow["url"]);
                            resultRow.Add("单次招股的说明书个数", matchedCount);
                            resultEW.AddRow(resultRow);
                        }
                    }
                }
            }
            resultEW.SaveToDisk();
        } 
    }
}