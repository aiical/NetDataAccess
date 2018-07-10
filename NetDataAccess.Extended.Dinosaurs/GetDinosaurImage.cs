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

namespace NetDataAccess.Extended.Dinosaurs
{
    public class GetDinosaurImage : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            this.GetImage(listSheet); 
            return true;
        }

        private void GetImage(IListSheet listSheet)
        {
            string[] ps = this.Parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            string imgDir = ps[0];


            String exportDir = this.RunPage.GetExportDir(); 
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir(); 
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string name = row["name"];
                string detailUrl = row["detailPageUrl"];
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string localFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir);
                    string imageFilePath = Path.Combine(imgDir, name + ".jpg");
                    File.Copy(localFilePath, imageFilePath);
                }
            } 
        }

         
    }
}