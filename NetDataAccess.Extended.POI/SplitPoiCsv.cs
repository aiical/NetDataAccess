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

namespace NetDataAccess.Extended.POI
{
    public class SplitPoiCsv : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            this.Split();
            return true;
        }

        private void Split()
        {
            string[] parameters = this.Parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            string sourceFilePath = parameters[0];
            string destDir = parameters[1];
            string filePrefix = parameters[2];
            int oneFileRowCount = int.Parse(parameters[3]);

            CsvSpliter cs = new CsvSpliter();
            int totalRowCount = cs.Init(sourceFilePath);

            int fileCount = (int)Math.Ceiling((double)totalRowCount / (double)oneFileRowCount);
            this.RunPage.InvokeAppendLogText("共需拆分成 " + fileCount + " 个文件", LogLevelType.System, true);
            for (int i = 0; i < fileCount; i++)
            {
                int fromIndex = oneFileRowCount * i;
                string destFilePath = Path.Combine(destDir, filePrefix + "_" + (i + 1).ToString().PadLeft(5, '0'));
                cs.GetPart(destFilePath, fromIndex, oneFileRowCount);
                this.RunPage.InvokeAppendLogText("已输出 FilePath = " + destFilePath, LogLevelType.System, true);
            }
            this.RunPage.InvokeAppendLogText("完成拆分.", LogLevelType.System, true);
        }
    }
}