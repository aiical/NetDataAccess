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
using System.Globalization;

namespace NetDataAccess.Extended.GuPiao
{
    /// <summary>
    /// 下载公告详情文件后，把pdf、html、txt文件等转换成txt格式并存储
    /// </summary>
    public class GetHuShenGongGaoDetailTxtFile : ExternalRunWebPage
    {

        private string _GongGaoSourceFileDir = null;
        private string GongGaoSourceFileDir
        {
            get
            {
                return this._GongGaoSourceFileDir;
            }
            set
            {
                this._GongGaoSourceFileDir = value;
            }
        }

        public override bool BeforeAllGrab()
        {
            string[] ps = this.Parameters.Split(new string[] { ","},  StringSplitOptions.RemoveEmptyEntries);
            this.GongGaoSourceFileDir = ps[0];
            return base.BeforeAllGrab();
        }

        public override void GetDataByOtherAcessType(Dictionary<string, string> listRow)
        {
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            base.GetDataByOtherAcessType(listRow);
            string detailUrl = listRow["detailPageUrl"];
            string adjunctType = listRow["adjunctType"].ToLower().Trim();
            string destFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir);
            string sourceFilePath = this.RunPage.GetFilePath(detailUrl, this.GongGaoSourceFileDir);

            try
            {
                switch (adjunctType)
                {
                    case "pdf":
                        Pdf2Txt.Pdf2TxtByITextSharp(sourceFilePath, destFilePath, true);
                        break;
                    case "html":
                        Html2Txt.Html2TxtByHtmlAgilityPack(sourceFilePath, destFilePath, true, "gb2312");
                        break;
                    case "txt":
                        {
                            Html2Txt.Html2TxtByHtmlAgilityPack(sourceFilePath, destFilePath, true, "gb2312");
                            //File.Copy(sourceFilePath, destFilePath);
                        }
                        break;
                    default:
                        Html2Txt.Html2TxtByHtmlAgilityPack(sourceFilePath, destFilePath, true, "gb2312");
                        break;
                        //throw new Exception("不可识别的公告文档类型, adjunctType = " + adjunctType);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        } 
    }
}