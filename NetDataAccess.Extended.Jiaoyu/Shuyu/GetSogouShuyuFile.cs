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

namespace NetDataAccess.Extended.Jiaoyu.Shuyu
{
    public class GetSogouShuyuFile : ExternalRunWebPage
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

        private void GetList(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("cate1", 0);
            resultColumnDic.Add("cateId1", 1);
            resultColumnDic.Add("cate2", 2);
            resultColumnDic.Add("cateId2", 3);
            resultColumnDic.Add("cate3", 4);
            resultColumnDic.Add("cateId3", 5);
            resultColumnDic.Add("name", 6);
            string resultFilePath = Path.Combine(exportDir, "教育_术语_scel文件信息.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null); 

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"];

                int fileIdBeginIndex = detailUrl.IndexOf("id=");
                int fileIdEndIndex = detailUrl.IndexOf("&");
                string fileId = detailUrl.Substring(fileIdBeginIndex + 3, fileIdEndIndex - fileIdBeginIndex - 3);

                string name = row["name"] + "_" + fileId + ".scel";
                string fromFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir);

                string fileDir = Path.Combine(Path.GetDirectoryName(pageSourceDir), "File");
                string toFilePath = this.RunPage.GetFilePath(name, fileDir);
                File.Copy(fromFilePath, toFilePath, true);
                
                string fileName = CommonUtil.ProcessFileName(name, "_") ;
                
                Dictionary<string, string> f2vs = new Dictionary<string, string>(); 
                f2vs.Add("cate1", row["cate1"]);
                f2vs.Add("cateId1", row["cateId1"]);
                f2vs.Add("cate2", row["cate2"]);
                f2vs.Add("cateId2", row["cateId2"]);
                f2vs.Add("cate3", row["cate3"]);
                f2vs.Add("cateId3", row["cateId3"]);
                f2vs.Add("name", fileName);
                resultEW.AddRow(f2vs);
            }
            resultEW.SaveToDisk();
        }
    }
}