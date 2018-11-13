using NetDataAccess.Base.Common;
using NetDataAccess.Base.Config;
using NetDataAccess.Base.UI;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;

namespace NetDataAccess.Extended.OCR.Tess
{
    public class GetTextFromImage : ExternalRunWebPage
    {
        private string _SourceImageDir = "";
        private string _TessractDataDir = "";
        private string _Language = "";
        private Dictionary<string, string> _TressractVariables = null;

        public override bool BeforeAllGrab()
        {
            string[] parameters = this.Parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            this._SourceImageDir = parameters[0];
            this._TessractDataDir = parameters[1];
            this._Language = parameters[2];

            this._TressractVariables = new Dictionary<string, string>();
            //this._TressractVariables.Add("tessedit_char_whitelist", "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ");

            return base.BeforeAllGrab();
        }

        public override void GetDataByOtherAccessType(Dictionary<string, string> listRow)
        {
            string imageName = listRow["imageName"];
            string detailPageUrl = listRow[SysConfig.DetailPageUrlFieldName];
            string sourceTextDir = this.RunPage.GetDetailSourceFileDir();
            string sourceTextFilePath = this.RunPage.GetFilePath(detailPageUrl, sourceTextDir);
            try
            {
                if (!File.Exists(sourceTextFilePath))
                {
                    string sourceImageFilePath = Path.Combine(this._SourceImageDir, imageName);
                    Bitmap bmp = new Bitmap(sourceImageFilePath);
                    string text = SimpleOCR.OCRMultiLine(bmp, this._TessractDataDir, this._Language, this._TressractVariables);

                    CommonUtil.CreateFileDirectory(sourceTextFilePath);
                    FileHelper.SaveTextToFile(text, sourceTextFilePath, Encoding.UTF8);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("识别图片出错, imageName = " + imageName, ex);
                throw ex;
            }
        }
    }
}
