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
using NetDataAccess.Base.Reader;

namespace NetDataAccess.Extended.Celloud
{
    public class OalibGetPaperHtmlPage : CustomProgramBase
    {
        public bool Run(string parameters, IListSheet listSheet)
        {
            //return GetAllListPageUrl(listSheet) && GetAllUrlHost(listSheet);
            return GetAllEachList(listSheet) && GetImportFile(listSheet);
            //return  GetImportFile(listSheet);
        }

        private bool GetImportFile(IListSheet listSheet)
        {
            string exportDir = this.RunPage.GetExportDir();
            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                "featuremanual",
                "genus",
                "genuscn", 
                "listfilepath",
                "searchsite",
                "sourcefilefolder",
                "species", 
                "txtfilefolder"});
            string resultFilePath = Path.Combine(exportDir, "论文列表.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);

            string speciesNameFilePath = Path.Combine(exportDir, "oalib初始化首页地址.xlsx");

            ExcelReader er = new ExcelReader(speciesNameFilePath);
            int rowCount = er.GetRowCount();
            for (int i = 0; i < rowCount; i++)
            {
                Dictionary<string, string> row = er.GetFieldValues(i); 
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string species = row["species"].Trim();
                    string listfilepath = "Celloud/oalib获取论文Html页/Export/oalib论文列表_" + species + ".xlsx";
                    string searchsite = "http://www.oalib.com";
                    string sourcefilefolder = "Celloud/oalib获取论文Html页/Detail/"; 
                    string txtfilefolder = "Celloud/oalib获取论文Html页/txt/";

                    Dictionary<string, string> f2vs = new Dictionary<string, string>(); 
                    f2vs.Add("listfilepath", listfilepath);
                    f2vs.Add("searchsite", searchsite);
                    f2vs.Add("sourcefilefolder", sourcefilefolder);
                    f2vs.Add("species", species); 
                    f2vs.Add("txtfilefolder", txtfilefolder); 
                    resultEW.AddRow(f2vs);
                }
            }
            er.Close();
            resultEW.SaveToDisk(); 
            return true;
        }

        private bool GetAllEachList(IListSheet listSheet)
        {
            string exportDir = this.RunPage.GetExportDir();
            string speciesNameFilePath = Path.Combine(exportDir, "oalib初始化首页地址.xlsx");

            ExcelReader er = new ExcelReader(speciesNameFilePath);
            int rowCount = er.GetRowCount();
            Dictionary<string, string> file2Types = this.GetFileTypes(listSheet);
            for (int i = 0; i < rowCount; i++)
            {
                Dictionary<string, string> row = er.GetFieldValues(i);
                string currentSpecies = row["species"].Trim();
                GetAllEachList(listSheet, file2Types, currentSpecies);
            }
            er.Close();
            return true;
        }

        private bool GetAllEachList(IListSheet listSheet, Dictionary<string, string> file2Types, string currentSpecies)
        {
            string exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                "abstractUrl",
                "url",
                "code", 
                "fileType",
                "filePath",
                "paperName",
                "species",
                "speciesCN",
                "genus",
                "genusCN",
                "publishYear"});
            string resultFilePath = Path.Combine(exportDir, "oalib论文列表_" + currentSpecies + ".xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;
            string detailPageNameColumnName = SysConfig.DetailPageNameFieldName;

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string url = row[detailPageUrlColumnName];
                    string code = row[detailPageNameColumnName];
                    string abstractUrl = row["abstractUrl"].Trim();
                    string species = row["species"].Trim();
                    if (species.ToLower() == currentSpecies.ToLower())
                    {  
                        string year = row["year"].Trim();
                        string paperName = row["paperName"];
                        string localFilePath = this.RunPage.GetFilePath(url, pageSourceDir);
                        string fileName = Path.GetFileName(localFilePath); 

                        try
                        {
                            string fileType = file2Types[url];
                            Dictionary<string, string> f2vs = new Dictionary<string, string>();
                            f2vs.Add("abstractUrl", abstractUrl);
                            f2vs.Add("url", url);
                            f2vs.Add("code", code);
                            f2vs.Add("fileType", fileType);
                            f2vs.Add("filePath", fileName);
                            f2vs.Add("paperName", paperName);
                            f2vs.Add("species", species); 
                            f2vs.Add("publishYear", year);
                            resultEW.AddRow(f2vs);
                        }
                        catch (Exception ex)
                        {
                            this.RunPage.InvokeAppendLogText("读取出错.  " + ex.Message + " LocalPath = " + localFilePath, LogLevelType.Error, true);
                            throw ex;
                        }
                    }
                }
            }
            resultEW.SaveToDisk();

            return true;
        }
        private Dictionary<string,string> GetFileTypes(IListSheet listSheet)
        {
            Dictionary<string, string> file2Types = new Dictionary<string, string>();

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;
            string detailPageNameColumnName = SysConfig.DetailPageNameFieldName;
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string url = row[detailPageUrlColumnName];
                    string code = row[detailPageNameColumnName];
                    string abstractUrl = row["abstractUrl"].Trim();
                    string species = row["species"].Trim();  
                    string year = row["year"].Trim();
                    string paperName = row["paperName"];
                    string localFilePath = this.RunPage.GetFilePath(url, pageSourceDir);
                    string fileName = Path.GetFileName(localFilePath);
                    FileStream fs = null;

                    try
                    {
                        if (!file2Types.ContainsKey(url))
                        {
                            fs = new FileStream(localFilePath, FileMode.Open);
                            string fileType = FileHelper.CheckFileType(fs);
                            file2Types.Add(url, fileType);
                        }
                    }
                    catch (Exception ex)
                    {
                        if (fs != null)
                        {
                            fs.Close();
                            fs.Dispose();
                            fs = null;
                        }
                        this.RunPage.InvokeAppendLogText("读取出错.  " + ex.Message + " LocalPath = " + localFilePath, LogLevelType.Error, true);
                        throw ex;
                    }
                }
            } 

            return file2Types;
        }


        private bool GetAllListPageUrl(IListSheet listSheet)
        {
            string exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                "abstractUrl",
                "url",
                "code", 
                "fileType",
                "filePath",
                "paperName",
                "species", 
                "publishYear"});
            string resultFilePath = Path.Combine(exportDir, "oalib论文列表.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;
            string detailPageNameColumnName = SysConfig.DetailPageNameFieldName;

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string url = row[detailPageUrlColumnName];
                    string code = row[detailPageNameColumnName];
                    string abstractUrl = row["abstractUrl"].Trim();
                    string species = row["species"].Trim(); 
                    string year = row["year"].Trim();
                    string paperName = row["paperName"];
                    string localFilePath = this.RunPage.GetFilePath(url, pageSourceDir);
                    string fileName = Path.GetFileName(localFilePath);
                    FileStream fs = null;

                    try
                    {
                        fs = new FileStream(localFilePath, FileMode.Open);
                        string fileType = FileHelper.CheckFileType(fs);
                        Dictionary<string, string> f2vs = new Dictionary<string, string>();
                        f2vs.Add("abstractUrl", abstractUrl);
                        f2vs.Add("url", url);
                        f2vs.Add("code", code);
                        f2vs.Add("fileType", fileType);
                        f2vs.Add("filePath", fileName);
                        f2vs.Add("paperName", paperName);
                        f2vs.Add("species", species); 
                        f2vs.Add("publishYear", year);
                        resultEW.AddRow(f2vs);
                    }
                    catch (Exception ex)
                    {
                        if (fs != null)
                        {
                            fs.Dispose();
                            fs = null;
                        }
                        this.RunPage.InvokeAppendLogText("读取出错.  " + ex.Message + " LocalPath = " + localFilePath, LogLevelType.Error, true);
                        throw ex;
                    }
                }
            }
            resultEW.SaveToDisk();

            return true;
        }

        private bool GetAllUrlHost(IListSheet listSheet)
        {
            string exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            
            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("hostName", 0);
            resultColumnDic.Add("fileType", 1);
            resultColumnDic.Add("canProcess", 2);
            resultColumnDic.Add("exampleUrl", 3);
            resultColumnDic.Add("paperContentXPath", 4);
            string resultFilePath = Path.Combine(exportDir, "oalib论文格式分类.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;
            string detailPageNameColumnName = SysConfig.DetailPageNameFieldName;
            List<string> allHosts= new  List<string>();
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string url = row[detailPageUrlColumnName];
                    Uri uri = new Uri(url);
                    string hostName = uri.Host;
                    if (!allHosts.Contains(hostName))
                    {
                        allHosts.Add(hostName);
                        Dictionary<string, string> f2vs = new Dictionary<string, string>();
                        f2vs.Add("hostName", hostName);
                        f2vs.Add("exampleUrl", url);
                        resultEW.AddRow(f2vs);
                    }
                }
            }
            resultEW.SaveToDisk();

            return true;
        }
    }
}