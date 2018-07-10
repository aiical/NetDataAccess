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
using NetDataAccess.Base.Writer;
using NetDataAccess.Base.Reader;
using NetDataAccess.Base.CsvHelper;
using NetDataAccess.Base.DB;
using HtmlAgilityPack;

namespace NetDataAccess.Extended.Yiguo
{
    /// <summary>
    /// 本来生活
    /// 从本地json文件中获取库存信息
    /// </summary>
    public class BenlaiDetailPageInfoWH : CustomProgramBase
    {
        #region 入口函数
        public bool Run(string parameters, IListSheet listSheet)
        {
            return this.GenerateDetailPageInfo(listSheet);
        }
        #endregion

        #region 生成库存信息文件，包含了之前获取到的商品其他信息
        private bool GenerateDetailPageInfo(IListSheet listSheet)
        {
            bool succeed = true;
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                "商品名称",
                "价格",
                "一级分类编码", 
                "一级分类", 
                "二级分类编码", 
                "二级分类",
                "三级分类编码", 
                "三级分类",
                "规格", 
                "温馨提示", 
                "满减", 
                "评论数", 
                "好评", 
                "中评", 
                "差评", 
                "好评度", 
                "url", 
                "地区", 
                "productPromotionWord",
                "原价",
                "商品编码",
                "是否有货"});


            Dictionary<string, string> resultColumnFormat = new Dictionary<string, string>();
            resultColumnFormat.Add("价格", "#,##0.00");
            resultColumnFormat.Add("评论数", "#,##0");
            resultColumnFormat.Add("好评", "#,##0");
            resultColumnFormat.Add("中评", "#,##0");
            resultColumnFormat.Add("差评", "#,##0");
            resultColumnFormat.Add("好评度", "0.00%");
            resultColumnFormat.Add("原价", "#,##0.00");

            string readDetailDir = this.RunPage.GetReadFileDir();
            string exportDir = this.RunPage.GetExportDir();
            string resultFilePath = Path.Combine(exportDir, "本来生活商品详情" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, resultColumnFormat);

            GenerateDetailPageInfo(listSheet, pageSourceDir, resultEW); 

            resultEW.SaveToDisk(); 

            return succeed;
        }
        #endregion

        #region 从json中获取库存信息文件，并保存之，包含了之前获取到的商品其他信息
        private void GenerateDetailPageInfo(IListSheet listSheet, string pageSourceDir, ExcelWriter resultEW)
        {
            try
            { 
                for (int i = 0; i < listSheet.RowCount; i++)
                {
                    Dictionary<string, string> row = listSheet.GetRow(i);
                    string pageUrl = listSheet.PageUrlList[i];
                    string pageName = listSheet.PageNameList[i];
                    string localFilePath = this.RunPage.GetFilePath(pageUrl, pageSourceDir);
                    string has = "";

                    TextReader tr = null;

                    try
                    {
                        tr = new StreamReader(localFilePath);
                        string webPageJson = tr.ReadToEnd();
                        JObject rootJo = JObject.Parse(webPageJson);
                        has = (rootJo["VisType"].ToString() == "0" && rootJo["Status"].ToString() == "1") ? "是" : "否";
                    }
                    catch (Exception ex)
                    {
                        this.RunPage.InvokeAppendLogText("读取出错. url = " + pageUrl + ". " + ex.Message, LogLevelType.Error, true);
                        throw ex;
                    }

                    Dictionary<string, object> f2vs = new Dictionary<string, object>();
                    f2vs.Add("是否有货", has);
                    f2vs.Add("商品名称", row["商品名称"]);
                    f2vs.Add("价格", decimal.Parse(row["价格"]));
                    f2vs.Add("一级分类编码", row["一级分类编码"]);
                    f2vs.Add("一级分类", row["一级分类"]);
                    f2vs.Add("二级分类编码", row["二级分类编码"]);
                    f2vs.Add("二级分类", row["二级分类"]);
                    f2vs.Add("三级分类编码", row["三级分类编码"]);
                    f2vs.Add("三级分类", row["三级分类"]);  
                    f2vs.Add("规格", row["规格"]);
                    f2vs.Add("温馨提示", row["温馨提示"]);
                    f2vs.Add("满减", row["满减"]);
                    f2vs.Add("评论数", int.Parse(row["评论数"]));
                    f2vs.Add("好评", int.Parse(row["好评"]));
                    f2vs.Add("中评", int.Parse(row["中评"]));
                    f2vs.Add("差评", int.Parse(row["差评"]));
                    if (!CommonUtil.IsNullOrBlank(row["好评度"]))
                    {
                        f2vs.Add("好评度", decimal.Parse(row["好评度"]));
                    }
                    f2vs.Add("url", row["url"]);
                    f2vs.Add("地区", row["地区"]);
                    f2vs.Add("productPromotionWord", row["商品编码"]);
                    if (!CommonUtil.IsNullOrBlank(row["原价"]))
                    {
                        f2vs.Add("原价", decimal.Parse(row["原价"]));
                    }
                    f2vs.Add("商品编码", row["商品编码"]);

                    resultEW.AddRow(f2vs);
                }
            }
            catch (Exception ex)
            { 
                throw ex;
            }
        }
        #endregion
    }
}