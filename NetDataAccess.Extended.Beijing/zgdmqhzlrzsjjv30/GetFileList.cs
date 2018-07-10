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

namespace NetDataAccess.Extended.Beijing.zgdmqhzlrzsjjv30
{
    /// <summary>
    /// 中国地面气候资料日值数据集(V3.0)，从txt中获取文件列表
    /// </summary>
    public class GetFileList : CustomProgramBase
    {
        #region 入口函数
        /// <summary>
        /// 入口函数
        /// </summary>
        /// <param name="parameters"></param>
        /// <param name="listSheet"></param>
        /// <returns></returns>
        public bool Run(string parameters, IListSheet listSheet)
        {
            return GetAllListPageUrl(parameters, listSheet);
        }
        #endregion

        #region 获取待下载的文件列表
        private bool GetAllListPageUrl(string parameters, IListSheet listSheet)
        { 
            //已经下载下来的首页html保存到的目录（文件夹）
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            string[] parameterArray = parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            string txtDirPath = parameterArray[0];
            string[] allTxtFilePaths = Directory.GetFiles(txtDirPath);
            
            //输出excel表格包含的列
            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                "detailPageUrl",
                "detailPageName", 
                "cookie",
                "grabStatus", 
                "giveUpGrab",
                "feature",
                "month"});

            //输出目录（从配置中获取）
            string exportDir = parameterArray[1];

            //输出文件的本地路径，此输出文件是项目“车维修获取列表页”的输入文件
            string resultFilePath = Path.Combine(exportDir, "获取数据文件_中国地面气候资料日值数据集V30.xlsx");

            //输出对象
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic);

            foreach (string txtFilePath in allTxtFilePaths)
            {
                StreamReader sr =null;
                try
                {
                    sr = new StreamReader(txtFilePath, Encoding.Default);
                    String line;
                    while ((line = sr.ReadLine()) != null)
                    {
                        //http://101.201.177.119/dlcdc/space/gpfs01/sdb/sdb_files/datasets/SURF_CLI_CHN_MUL_DAY_V3.0/datasets/SSD/SURF_CLI_CHN_MUL_DAY-SSD-14032-201001.TXT?Expires=1463721011&OSSAccessKeyId=CcULE6lAfEbIFtKD&Signature=ylFxbnK9kMmel0H8xHwglfEEtXI%3D
                        string dataFilePath = line;
                        int questionMarkIndex = dataFilePath.IndexOf("?");
                        int lastSlashIndex = dataFilePath.Substring(0, questionMarkIndex).LastIndexOf("/");
                        int previousSlashIndex = dataFilePath.Substring(0, lastSlashIndex).LastIndexOf("/");
                        string feature = dataFilePath.Substring(previousSlashIndex + 1, lastSlashIndex - previousSlashIndex-1);
                        string fileName = dataFilePath.Substring(lastSlashIndex + 1, questionMarkIndex - lastSlashIndex-1);
                        string month = fileName.Substring(fileName.Length - 10, 6);
                        Dictionary<string, object> f2vs = new Dictionary<string, object>();
                        f2vs.Add("detailPageUrl", dataFilePath);
                        f2vs.Add("detailPageName", fileName);
                        f2vs.Add("feature", feature);
                        f2vs.Add("month", month);
                        resultEW.AddRow(f2vs);
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception("读取列表文件出错." + txtFilePath, ex);
                }
            }
                 

            //保存到硬盘
            resultEW.SaveToDisk();

            return true;
        }
        #endregion 
    }
}