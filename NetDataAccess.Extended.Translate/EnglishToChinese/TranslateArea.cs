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

namespace NetDataAccess.Extended.Translate.EnglishToChinese
{
    /// <summary>
    /// 翻译国家或者地区
    /// </summary>
    public class TranslateArea : ExternalRunWebPage
    { 
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
        public override bool BeforeAllGrab()
        {
            this.LoadDictionary();
            return base.BeforeAllGrab();
        }

        public override void GetDataByOtherAcessType(Dictionary<string, string> listRow)
        { 
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            string fromNameStr = listRow["fromName"]; 
            List<string> toNameList = new List<string>();
            List<string> fromNameList = new List<string>();

            string[] fromNames = fromNameStr.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < fromNames.Length; i++)
            {
                string fromName = fromNames[i].Trim();
                string fromNameLow = fromName.ToLower();
                if (this.Dic.ContainsKey(fromNameLow))
                {
                    string toName = this.Dic[fromNameLow];
                    toNameList.Add(toName);

                    fromNameList.Add(fromName);
                }
                else
                {
                    throw new Exception("无法翻译, fromName = " + fromName);
                }
            }

            CsvWriter tempCsvWriter = this.GetCsvWriter(listRow);
            Dictionary<string, string> row = new Dictionary<string, string>();
            row.Add("fromName", CommonUtil.StringArrayToString(fromNameList.ToArray(), ", "));
            row.Add("toName", CommonUtil.StringArrayToString(toNameList.ToArray(), ", "));
            tempCsvWriter.AddRow(row);
            tempCsvWriter.SaveToDisk();            
        }


        private CsvWriter GetCsvWriter(Dictionary<string, string> listRow)
        {
            string detailUrl = listRow["detailPageUrl"];
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            string resultTextFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir);
            CommonUtil.CreateFileDirectory(resultTextFilePath);

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("fromName", 0);
            resultColumnDic.Add("toName", 1);
            CsvWriter resultEW = new CsvWriter(resultTextFilePath, resultColumnDic);
            return resultEW;
        }


        private void GetList(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("fromName", 0);
            resultColumnDic.Add("toName", 1);
            string resultFilePath = Path.Combine(exportDir, "翻译结果.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null); 
            
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"];
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    try
                    {
                        string resultTextFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir);

                        CsvReader csvReader = new CsvReader(resultTextFilePath);
                        Dictionary<string, string> f2vs = csvReader.GetFieldValues(0);  
                        resultEW.AddRow(f2vs);
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    } 
                }
            }
            resultEW.SaveToDisk();
        }

        private Dictionary<string, string> _Dic = new Dictionary<string, string>();
        private Dictionary<string, string> Dic
        {
            get
            {
                return this._Dic;
            }
        }

        private void LoadDictionary()
        {
            string[] parameters = this.Parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            string dicFilePath = parameters[0];
            ExcelReader er = new ExcelReader(dicFilePath, "List");
            Dictionary<string, string> dicValues = new Dictionary<string, string>();
            int rowCount = er.GetRowCount();
            for (int i = 0; i < rowCount; i++)
            {
                Dictionary<string, string> row = er.GetFieldValues(i);
                string cnName = row["中文名"];
                string enName = row["英文名"];
                this.InsertToDic(enName, cnName, dicValues);

                string sxName = row["缩写"];
                this.InsertToDic(enName, cnName, dicValues);

                string qtywmName = row["其它英文名"];
                if (qtywmName.Length != 0)
                {
                    string[] qtywmNames = qtywmName.Split(new string[] { ";" }, StringSplitOptions.RemoveEmptyEntries);
                    for (int j = 0; j < qtywmNames.Length; j++)
                    {
                        this.InsertToDic(qtywmNames[j], cnName, dicValues);
                    }
                }
            }
            this._Dic = dicValues;
        }
        private void InsertToDic(string fromName, string toName, Dictionary<string, string> dic)
        {
            fromName = fromName.ToLower().Trim();
            if (fromName.Length != 0 && !dic.ContainsKey(fromName))
            {
                dic.Add(fromName.Trim(), toName.Trim());
            }
        }
    }
}