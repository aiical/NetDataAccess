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

namespace NetDataAccess.Extended.Ez4s
{
    /// <summary>
    /// 车维修获取详情页
    /// </summary>
    public class CheweixiuGetDetailPage : CustomProgramBase
    {
        #region 入口函数
        /// <summary>
        /// 入口函数
        /// </summary>
        /// <param name="parameters">“扩展程序配置”信息中的parameters属性值</param>
        /// <param name="listSheet">输入文件，记录了要下载的所有页面（其中包含的记录，再爬取前先由excel导入到sqlite表，然后系统操作的一直是sqlite表里的记录）</param>
        /// <returns></returns>
        public bool Run(string parameters, IListSheet listSheet)
        {
            return GetAllPageDetailInfo(listSheet);
        }
        #endregion

        #region 逐个详情页获取信息
        /// <summary>
        /// 逐个详情页获取信息
        /// </summary>
        /// <param name="listSheet"></param>
        /// <returns></returns>
        private bool GetAllPageDetailInfo(IListSheet listSheet)
        {
            //输出目录（从配置中获取）
            string exportDir = this.RunPage.GetExportDir();

            //下载下来的详情页html保存目录（文件夹）
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            //输出excel表格包含的列，此文件提供给客户
            Dictionary<string, int> resultColumnDic = CommonUtil.InitStringIndexDic(new string[]{
                "名称", 
                "详细地址",
                "联系电话",
                "lng",
                "lat", 
                "地段", 
                "地标建筑",
                "营业时间", 
                "服务站类型", 
                "业务", 
                "擅长品牌", 
                "店家资质", 
                "占地面积", 
                "技师工人", 
                "工位数量", 
                "设备信息", 
                "工时费", 
                "服务", 
                "周边环境", 
                "店家简介",
                "url" });

            //输出文件地址
            string resultFilePath = Path.Combine(exportDir, "车维修维修站点详情.xlsx");

            //输出文件对象
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;
            string detailPageNameColumnName = SysConfig.DetailPageNameFieldName;

            //循环输入文件中的所有行
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                //获取输入文件中此行的url及属性
                Dictionary<string, string> row = listSheet.GetRow(i);

                //如果此行没有放弃爬取（爬取工具可以配置成爬取失败后放弃爬取，被放弃爬取的行记录到sqlite表中）
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string url = row[detailPageUrlColumnName];

                    //"名称"
                    string name = row["name"].Trim();

                    //经纬度
                    Nullable<decimal> lng = null;
                    Nullable<decimal> lat = null;
                    
                    //"详细地址"
                    string xxdz = "";

                    //"联系电话",
                    string lxdh = "";

                    //"地段", 
                    string dd = "";

                    //"地标建筑",
                    string dbjz = "";

                    //"营业时间", 
                    string yysj = "";

                    //"服务站类型", 
                    string fwzlx = "";

                    //"业务", 
                    string yw = "";

                    //"擅长品牌", 
                    string scpp = "";

                    //"店家资质", 
                    string djzz = "";

                    //"占地面积", 
                    string zdmj = "";

                    //"技师工人", 
                    string jsgr = "";

                    //"工位数量", 
                    string gwsl = "";

                    //"设备信息", 
                    string sbxx = "";

                    //"工时费", 
                    string gsf = "";

                    //"服务", 
                    string fw = "";

                    //"周边环境", 
                    string zbhj = "";

                    //"店家简介", 
                    string djjj = "";


                    //下载下来的html地址
                    string localFilePath = this.RunPage.GetFilePath(url, pageSourceDir);
                    TextReader tr = null;

                    try
                    {
                        //读取展会详情页html，并加载到HtmlDocument对象中（系统提供了构造HtmlDocument对象的方法，写这段代码的时候忘记用了）
                        tr = new StreamReader(localFilePath);
                        string webPageHtml = tr.ReadToEnd();
                        HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                        htmlDoc.LoadHtml(webPageHtml);

                        HtmlNodeCollection allGroupAPropertyNodes = htmlDoc.DocumentNode.SelectNodes("//div[@class=\"sdesc clearfix\"]/div[@class=\"left\"]/table/tr");
                        foreach (HtmlNode propertyNode in allGroupAPropertyNodes)
                        {
                            HtmlNodeCollection nvNodes = propertyNode.SelectNodes("./td");
                            if (nvNodes.Count > 0)
                            {
                                HtmlNode nameNode = nvNodes[0];
                                HtmlNode valueNode = nvNodes.Count > 1 ? nvNodes[1] : null;
                                string propertyName = CommonUtil.HtmlDecode(nameNode.InnerText).Trim();
                                string propertyValue = valueNode == null ? "" : CommonUtil.HtmlDecode(valueNode.InnerText).Trim();
                                switch (propertyName)
                                {
                                    case "详细地址":
                                        xxdz = propertyValue;
                                        break;
                                    case "联系电话":
                                        lxdh = propertyValue;
                                        break;
                                    case "地段":
                                        dd = propertyValue;
                                        break;
                                    case "地标建筑":
                                        dbjz = propertyValue;
                                        break;
                                    case "营业时间":
                                        yysj = propertyValue;
                                        break;
                                }
                            }
                        }

                        HtmlNodeCollection allGroupBPropertyNodes = htmlDoc.DocumentNode.SelectNodes("//dl[@id=\"sinfos\"]/dt");
                        foreach (HtmlNode nameNode in allGroupBPropertyNodes)
                        {
                            HtmlNode nextNode = nameNode.NextSibling;
                            HtmlNode valueNode =null;
                            while (nextNode != null)
                            {
                                if (nextNode.Name.ToLower() == "dd")
                                {
                                    valueNode = nextNode;
                                    break;
                                }
                                else
                                {
                                    nextNode = nextNode.NextSibling;
                                }
                            }
                            if (valueNode != null)
                            {
                                string propertyName = CommonUtil.HtmlDecode(nameNode.InnerText).Trim();
                                switch (propertyName)
                                {
                                    case "业务":
                                        {
                                            HtmlNodeCollection allPNodes = valueNode.SelectNodes("./p");
                                            foreach (HtmlNode pNode in allPNodes)
                                            {
                                                string pValue = CommonUtil.HtmlDecode(pNode.InnerText).Trim();
                                                if (pValue.StartsWith("店铺类型"))
                                                {
                                                    //服务站类型
                                                    HtmlNode fwzlxNode = pNode.SelectSingleNode("./a");
                                                    fwzlx = CommonUtil.HtmlDecode(fwzlxNode.InnerText).Trim();
                                                }
                                                else
                                                {
                                                    yw += (pValue + "\r\n");
                                                }
                                            }
                                            yw = yw.Trim();
                                        }
                                        break;
                                    case "擅长品牌":
                                        {
                                            scpp = CommonUtil.HtmlDecode(valueNode.InnerText).Trim();
                                        }
                                        break;
                                    case "店家资质":
                                        {
                                            djzz = CommonUtil.HtmlDecode(valueNode.InnerText).Trim();
                                        }
                                        break;
                                    case "占地面积":
                                        {
                                            zdmj = CommonUtil.HtmlDecode(valueNode.InnerText).Trim();
                                        }
                                        break;
                                    case "技师工人":
                                        {
                                            jsgr = CommonUtil.HtmlDecode(valueNode.InnerText).Trim();
                                        }
                                        break;
                                    case "工位数量":
                                        {
                                            gwsl = CommonUtil.HtmlDecode(valueNode.InnerText).Trim();
                                        }
                                        break;
                                    case "设备信息":
                                        {
                                            HtmlNodeCollection allPNodes = valueNode.SelectNodes("./p");
                                            foreach (HtmlNode pNode in allPNodes)
                                            {
                                                string pValue = CommonUtil.HtmlDecode(pNode.InnerText).Trim();
                                                sbxx += (pValue + "\r\n");
                                            }
                                            sbxx = sbxx.Trim();
                                        }
                                        break;
                                    case "工时费":
                                        {
                                            gsf = CommonUtil.HtmlDecode(valueNode.InnerText).Trim();
                                        }
                                        break;
                                    case "服务":
                                        {
                                            fw = CommonUtil.HtmlDecode(valueNode.InnerText).Trim();
                                        }
                                        break;
                                    case "周边环境":
                                        {
                                            zbhj = CommonUtil.HtmlDecode(valueNode.InnerText).Trim();
                                        }
                                        break;
                                    case "店家简介":
                                        {
                                            djjj = CommonUtil.HtmlDecode(valueNode.InnerText).Trim();
                                        }
                                        break;
                                }
                            }
                        }

                        HtmlNodeCollection allScriptNodes = htmlDoc.DocumentNode.SelectNodes("//script");
                        if (allScriptNodes != null)
                        {
                            foreach (HtmlNode scriptNode in allScriptNodes)
                            {
                                string script = scriptNode.InnerText;
                                if (script.Contains("var lng = "))
                                {
                                    int lngBeginIndex = script.IndexOf("var lng = ") + 10;
                                    int lngEndIndex = script.IndexOf(";", lngBeginIndex);
                                    int latBeginIndex = script.IndexOf("var lat = ") + 10;
                                    int latEndIndex = script.IndexOf(";", latBeginIndex);
                                    lng = decimal.Parse(script.Substring(lngBeginIndex, lngEndIndex - lngBeginIndex));
                                    lat = decimal.Parse(script.Substring(latBeginIndex, latEndIndex - latBeginIndex));
                                    break;
                                }
                            }
                        }

                        //保存信息到resultEW（ExcelWriter对象）
                        Dictionary<string, object> f2vs = new Dictionary<string, object>();
                        f2vs.Add("名称", name);
                        f2vs.Add("详细地址", xxdz);
                        f2vs.Add("联系电话", lxdh);
                        f2vs.Add("地段", dd);
                        f2vs.Add("地标建筑", dbjz);
                        f2vs.Add("营业时间", yysj);
                        f2vs.Add("服务站类型", fwzlx);
                        f2vs.Add("业务", yw);
                        f2vs.Add("擅长品牌", scpp);
                        f2vs.Add("店家资质", djzz);
                        f2vs.Add("占地面积", zdmj);
                        f2vs.Add("技师工人", jsgr);
                        f2vs.Add("工位数量", gwsl);
                        f2vs.Add("设备信息", sbxx);
                        f2vs.Add("工时费", gsf);
                        f2vs.Add("服务", fw);
                        f2vs.Add("周边环境", zbhj);
                        f2vs.Add("店家简介", djjj);
                        f2vs.Add("lng", lng);
                        f2vs.Add("lat", lat);
                        f2vs.Add("url", url);
                        resultEW.AddRow(f2vs);
                    }
                    catch (Exception ex)
                    {
                        if (tr != null)
                        {
                            tr.Dispose();
                            tr = null;
                        }
                        this.RunPage.InvokeAppendLogText("读取出错.  " + ex.Message + " LocalPath = " + localFilePath, LogLevelType.Error, true);
                        throw ex;
                    }
                }
            }

            //输出到硬盘
            resultEW.SaveToDisk();

            return true;
        }
        #endregion 
    }
}