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

namespace NetDataAccess.Extended.Tanyin
{
    /// <summary>
    /// 节能搜索低碳获取图片
    /// 即，获取低碳相关文章中信息正文中包含的图片
    /// 运行此程序前，系统已经提前爬取了listSheet中指定的图片
    /// 现在运行此扩展程序，记录下来所有图片的url和本地url的对照
    /// </summary>
    public class JnssdtGetImage : CustomProgramBase
    {
        #region 入口函数
        /// <summary>
        /// 入口函数
        /// </summary>
        /// <param name="parameters">“扩展程序配置”信息中的parameters属性值</param>
        /// <param name="listSheet">输入文件，记录了要下载的所有页面</param>
        /// <returns></returns>
        public bool Run(string parameters, IListSheet listSheet)
        {
            return SaveAllPageImageUrlToFile(listSheet);
        }
        #endregion

        #region 记录所有图片URL
        private bool SaveAllPageImageUrlToFile(IListSheet listSheet)
        { 
            //输出目录（从配置中获取）
            string exportDir = this.RunPage.GetExportDir();

            //下载下来的File的保存目录（文件夹）
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            //输出excel表格包含的列，此文件提供给客户
            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                "文章网址", 
                "图片网址",
                "下载成功", 
                "文件名"});

            //输出文件地址
            string resultFilePath = Path.Combine(exportDir, "新产品新技术-节能搜索低碳(图片).xlsx");

            //输出文件对象
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;
            string detailPageNameColumnName = SysConfig.DetailPageNameFieldName;

            //循环输入文件中的所有行
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                string url = row[detailPageUrlColumnName];
                string pageCode = row["pageCode"];
                string localFilePath = this.RunPage.GetFilePath(url, pageSourceDir);
                string fileName = Path.GetFileName(localFilePath);

                Dictionary<string, object> f2vs = new Dictionary<string, object>();
                f2vs.Add("文章网址", pageCode);
                f2vs.Add("图片网址", url);
                f2vs.Add("下载成功", giveUp ? "否" : "是");
                f2vs.Add("文件名", fileName); 
                resultEW.AddRow(f2vs);
            }

            //输出到硬盘
            resultEW.SaveToDisk();

            return true;
        }
        #endregion 
    }
}