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

namespace NetDataAccess.Extended.Jianzhu.QGJZSCJGGGFWPT_XM
{
    public class GetProjectPage : ExternalRunWebPage
    { 
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            try
            {
                return this.GetAllPages(listSheet);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private CsvWriter GetMainCsvWriter()
        {
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("项目编号", 0);
            resultColumnDic.Add("省级项目编号", 1);
            resultColumnDic.Add("项目名称", 2);
            resultColumnDic.Add("所在区划", 3);
            resultColumnDic.Add("建设单位", 4);
            resultColumnDic.Add("建设单位组织机构代码（统一社会信用代码）", 5);
            resultColumnDic.Add("项目分类", 6);
            resultColumnDic.Add("建设性质", 7);
            resultColumnDic.Add("工程用途", 8);
            resultColumnDic.Add("总投资", 9);
            resultColumnDic.Add("总面积", 10);
            resultColumnDic.Add("立项级别", 11);
            resultColumnDic.Add("立项文号", 12);

            string resultFilePath = Path.Combine(exportDir, "项目数据_项目信息.csv");
            CsvWriter resultEW = new CsvWriter(resultFilePath, resultColumnDic);
            return resultEW;
        }
        private CsvWriter GetZtbCsvWriter()
        {
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("项目编码", 0);
            resultColumnDic.Add("招标类型", 1);
            resultColumnDic.Add("招标方式", 2);
            resultColumnDic.Add("中标单位名称", 3);
            resultColumnDic.Add("中标日期", 4);
            resultColumnDic.Add("中标金额（万元）", 5);
            resultColumnDic.Add("中标通知书编号", 6);
            resultColumnDic.Add("省级中标通知书编号", 7);

            string resultFilePath = Path.Combine(exportDir, "项目数据_项目信息_招投标.csv");
            CsvWriter resultEW = new CsvWriter(resultFilePath, resultColumnDic);
            return resultEW;
        }
        private CsvWriter GetSgtscCsvWriter()
        {
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("项目编码", 0);
            resultColumnDic.Add("施工图审查合格书编号", 1);
            resultColumnDic.Add("省级施工图审查合格书编号", 2);
            resultColumnDic.Add("勘察单位名称", 3);
            resultColumnDic.Add("设计单位名称", 4);
            resultColumnDic.Add("施工图审查机构名称", 5);
            resultColumnDic.Add("审查完成日期", 6);

            string resultFilePath = Path.Combine(exportDir, "项目数据_项目信息_施工图审查.csv");
            CsvWriter resultEW = new CsvWriter(resultFilePath, resultColumnDic);
            return resultEW;
        }
        private CsvWriter GetHtbaCsvWriter()
        {
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("项目编码", 0);
            resultColumnDic.Add("合同类别", 1);
            resultColumnDic.Add("合同备案编号", 2);
            resultColumnDic.Add("省级合同备案编号", 3);
            resultColumnDic.Add("合同金额（万元）", 4);
            resultColumnDic.Add("合同签订日期", 5);

            string resultFilePath = Path.Combine(exportDir, "项目数据_项目信息_合同备案.csv");
            CsvWriter resultEW = new CsvWriter(resultFilePath, resultColumnDic);
            return resultEW;
        }
        private CsvWriter GetSgxkCsvWriter()
        {
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("项目编码", 0);
            resultColumnDic.Add("施工许可证编号", 1);
            resultColumnDic.Add("省级施工许可证编号", 2);
            resultColumnDic.Add("合同金额（万元）", 3);
            resultColumnDic.Add("面积（平方米）", 4);
            resultColumnDic.Add("发证日期", 5);

            string resultFilePath = Path.Combine(exportDir, "项目数据_项目信息_施工许可.csv");
            CsvWriter resultEW = new CsvWriter(resultFilePath, resultColumnDic);
            return resultEW;
        }
        private CsvWriter GetJgysbaCsvWriter()
        {
            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("项目编码", 0);
            resultColumnDic.Add("竣工备案编号", 1);
            resultColumnDic.Add("省级竣工备案编号", 2);
            resultColumnDic.Add("实际造价（万元）", 3);
            resultColumnDic.Add("实际面积（平方米）", 4);
            resultColumnDic.Add("实际开工日期", 5);
            resultColumnDic.Add("实际竣工验收日期", 6);

            string resultFilePath = Path.Combine(exportDir, "项目数据_项目信息_竣工验收备案.csv");
            CsvWriter resultEW = new CsvWriter(resultFilePath, resultColumnDic);
            return resultEW;
        }

        private bool GetAllPages(IListSheet listSheet)
        {
            CsvWriter mainCW = this.GetMainCsvWriter();
            CsvWriter ztbCW = this.GetZtbCsvWriter();
            CsvWriter sgtscCW = this.GetSgtscCsvWriter();
            CsvWriter htbaCW = this.GetHtbaCsvWriter();
            CsvWriter sgxkCW = this.GetSgxkCsvWriter();
            CsvWriter jgysbaCW = this.GetJgysbaCsvWriter(); 
            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;
            Dictionary<string, string> projectDic = new Dictionary<string, string>();
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailPageUrl = row[SysConfig.DetailPageUrlFieldName];
                string detailPageName = row[SysConfig.DetailPageNameFieldName];
                try
                {

                    bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                    if (!giveUp)
                    {
                        HtmlAgilityPack.HtmlDocument pageHtmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);

                        #region 基础信息
                        string xmmc = "";
                        string xmbh = "";
                        string sjxmbh = "";
                        string szqh = "";
                        string jsdw = "";
                        string jsdwzzjgdm = "";
                        string xmfl = "";
                        string jsxz = "";
                        string gcyt = "";
                        string ztz = "";
                        string zmj = "";
                        string lxjb = "";
                        string lxwh = "";
                        HtmlNode xmmcNode = pageHtmlDoc.DocumentNode.SelectSingleNode("//div[@class=\"user_info spmtop\"]");
                        if (xmmcNode == null)
                        {
                            throw new Exception("没有找到项目名称节点");
                        }
                        else
                        {
                            xmmc = CommonUtil.HtmlDecode(xmmcNode.InnerText.Trim()).Trim();
                        }

                        HtmlNodeCollection projectFieldNodeList = pageHtmlDoc.DocumentNode.SelectNodes("//div[@class=\"query_info_box \"]/div/div[@class=\"activeTinyTabContent\"]/dl/dd");
                        if (projectFieldNodeList != null)
                        {
                            for (int j = 0; j < projectFieldNodeList.Count; j++)
                            {
                                HtmlNode projectFieldNode = projectFieldNodeList[j];
                                string fieldText = projectFieldNode.InnerText.Trim();
                                int sIndex = fieldText.IndexOf("：");
                                string fieldName = CommonUtil.HtmlDecode(fieldText.Substring(0, sIndex)).Trim();
                                string fieldValue = CommonUtil.HtmlDecode(fieldText.Substring(sIndex + 1)).Trim();
                                switch (fieldName)
                                {
                                    case "项目编号":
                                        xmbh = fieldValue;
                                        break;
                                    case "省级项目编号":
                                        sjxmbh = fieldValue;
                                        break;
                                    case "所在区划":
                                        szqh = fieldValue;
                                        break;
                                    case "建设单位":
                                        jsdw = fieldValue;
                                        break;
                                    case "建设单位组织机构代码（统一社会信用代码）":
                                        jsdwzzjgdm = fieldValue;
                                        break;
                                    case "项目分类":
                                        xmfl = fieldValue;
                                        break;
                                    case "建设性质":
                                        jsxz = fieldValue;
                                        break;
                                    case "工程用途":
                                        gcyt = fieldValue;
                                        break;
                                    case "总投资":
                                        ztz = fieldValue;
                                        break;
                                    case "总面积":
                                        zmj = fieldValue;
                                        break;
                                    case "立项级别":
                                        lxjb = fieldValue;
                                        break;
                                    case "立项文号":
                                        lxwh = fieldValue;
                                        break;
                                }

                            }
                        }
                        else
                        {
                            throw new Exception("无法获取项目基本信息属性值");
                        }

                        Dictionary<string, string> f2vs = new Dictionary<string, string>();
                        f2vs.Add("项目编号", xmbh);
                        f2vs.Add("省级项目编号", sjxmbh);
                        f2vs.Add("项目名称", xmmc);
                        f2vs.Add("所在区划", szqh);
                        f2vs.Add("建设单位", jsdw);
                        f2vs.Add("建设单位组织机构代码（统一社会信用代码）", jsdwzzjgdm);
                        f2vs.Add("项目分类", xmfl);
                        f2vs.Add("建设性质", jsxz);
                        f2vs.Add("工程用途", gcyt);
                        f2vs.Add("总投资", ztz);
                        f2vs.Add("总面积", zmj);
                        f2vs.Add("立项级别", lxjb);
                        f2vs.Add("立项文号", lxwh);
                        mainCW.AddRow(f2vs);
                        #endregion

                        #region 招投标
                        HtmlNodeCollection ztbNodeList = pageHtmlDoc.DocumentNode.SelectNodes("//div[@id=\"tab_ztb\"]/table/tbody/tr[@class=\"row\"]");
                        if (ztbNodeList != null)
                        {
                            foreach (HtmlNode ztbNode in ztbNodeList)
                            {
                                HtmlNodeCollection ztbFieldNodeList = ztbNode.SelectNodes("./td");
                                Dictionary<string, string> ztbF2vs = new Dictionary<string, string>();
                                ztbF2vs.Add("项目编码", xmbh);
                                ztbF2vs.Add("招标类型", CommonUtil.HtmlDecode(ztbFieldNodeList[1].InnerText.Trim()));
                                ztbF2vs.Add("招标方式", CommonUtil.HtmlDecode(ztbFieldNodeList[2].InnerText.Trim()));
                                ztbF2vs.Add("中标单位名称", CommonUtil.HtmlDecode(ztbFieldNodeList[3].InnerText.Trim()));
                                ztbF2vs.Add("中标日期", CommonUtil.HtmlDecode(ztbFieldNodeList[4].InnerText.Trim()));
                                ztbF2vs.Add("中标金额（万元）", CommonUtil.HtmlDecode(ztbFieldNodeList[5].InnerText.Trim()));
                                ztbF2vs.Add("中标通知书编号", CommonUtil.HtmlDecode(ztbFieldNodeList[6].InnerText.Trim()));
                                ztbF2vs.Add("省级中标通知书编号", CommonUtil.HtmlDecode(ztbFieldNodeList[7].InnerText.Trim()));
                                ztbCW.AddRow(ztbF2vs);
                            }
                        }
                        #endregion

                        #region 施工图审查
                        HtmlNodeCollection sgtscNodeList = pageHtmlDoc.DocumentNode.SelectNodes("//div[@id=\"tab_sgtsc\"]/table/tbody/tr[@class=\"row\"]");
                        if (sgtscNodeList != null)
                        {
                            foreach (HtmlNode sgtscNode in sgtscNodeList)
                            {
                                HtmlNodeCollection sgtscFieldNodeList = sgtscNode.SelectNodes("./td");
                                Dictionary<string, string> sgtscF2vs = new Dictionary<string, string>();
                                sgtscF2vs.Add("项目编码", xmbh);
                                sgtscF2vs.Add("施工图审查合格书编号", CommonUtil.HtmlDecode(sgtscFieldNodeList[1].InnerText.Trim()));
                                sgtscF2vs.Add("省级施工图审查合格书编号", CommonUtil.HtmlDecode(sgtscFieldNodeList[2].InnerText.Trim()));
                                sgtscF2vs.Add("勘察单位名称", CommonUtil.HtmlDecode(sgtscFieldNodeList[3].InnerText.Trim()));
                                sgtscF2vs.Add("设计单位名称", CommonUtil.HtmlDecode(sgtscFieldNodeList[4].InnerText.Trim()));
                                sgtscF2vs.Add("施工图审查机构名称", CommonUtil.HtmlDecode(sgtscFieldNodeList[5].InnerText.Trim()));
                                sgtscF2vs.Add("审查完成日期", CommonUtil.HtmlDecode(sgtscFieldNodeList[6].InnerText.Trim()));
                                sgtscCW.AddRow(sgtscF2vs);
                            }
                        }
                        #endregion

                        #region 合同备案
                        HtmlNodeCollection htbaNodeList = pageHtmlDoc.DocumentNode.SelectNodes("//div[@id=\"tab_htba\"]/table/tbody/tr[@class=\"row\"]");
                        if (htbaNodeList != null)
                        {
                            foreach (HtmlNode htbaNode in htbaNodeList)
                            {
                                HtmlNodeCollection htbaFieldNodeList = htbaNode.SelectNodes("./td");
                                Dictionary<string, string> htbaF2vs = new Dictionary<string, string>();
                                htbaF2vs.Add("项目编码", xmbh);
                                htbaF2vs.Add("合同类别", CommonUtil.HtmlDecode(htbaFieldNodeList[1].InnerText.Trim()));
                                htbaF2vs.Add("合同备案编号", CommonUtil.HtmlDecode(htbaFieldNodeList[2].InnerText.Trim()));
                                htbaF2vs.Add("省级合同备案编号", CommonUtil.HtmlDecode(htbaFieldNodeList[3].InnerText.Trim()));
                                htbaF2vs.Add("合同金额（万元）", CommonUtil.HtmlDecode(htbaFieldNodeList[4].InnerText.Trim()));
                                htbaF2vs.Add("合同签订日期", CommonUtil.HtmlDecode(htbaFieldNodeList[5].InnerText.Trim()));
                                htbaCW.AddRow(htbaF2vs);
                            }
                        }
                        #endregion

                        #region 施工许可
                        HtmlNodeCollection sgxkNodeList = pageHtmlDoc.DocumentNode.SelectNodes("//div[@id=\"tab_sgxk\"]/table/tbody/tr[@class=\"row\"]");
                        if (sgxkNodeList != null)
                        {
                            foreach (HtmlNode sgxkNode in sgxkNodeList)
                            {
                                HtmlNodeCollection sgxkFieldNodeList = sgxkNode.SelectNodes("./td");
                                Dictionary<string, string> sgxkF2vs = new Dictionary<string, string>();
                                sgxkF2vs.Add("项目编码", xmbh);
                                sgxkF2vs.Add("施工许可证编号", CommonUtil.HtmlDecode(sgxkFieldNodeList[1].InnerText.Trim()));
                                sgxkF2vs.Add("省级施工许可证编号", CommonUtil.HtmlDecode(sgxkFieldNodeList[2].InnerText.Trim()));
                                sgxkF2vs.Add("合同金额（万元）", CommonUtil.HtmlDecode(sgxkFieldNodeList[3].InnerText.Trim()));
                                sgxkF2vs.Add("面积（平方米）", CommonUtil.HtmlDecode(sgxkFieldNodeList[4].InnerText.Trim()));
                                sgxkF2vs.Add("发证日期", CommonUtil.HtmlDecode(sgxkFieldNodeList[5].InnerText.Trim()));
                                sgxkCW.AddRow(sgxkF2vs);
                            }
                        }
                        #endregion

                        #region 竣工验收备案
                        HtmlNodeCollection jgysbaNodeList = pageHtmlDoc.DocumentNode.SelectNodes("//div[@id=\"tab_jgysba\"]/table/tbody/tr[@class=\"row\"]");
                        if (jgysbaNodeList != null)
                        {
                            foreach (HtmlNode jgysbaNode in jgysbaNodeList)
                            {
                                HtmlNodeCollection jgysbaFieldNodeList = jgysbaNode.SelectNodes("./td");
                                Dictionary<string, string> jgysbaF2vs = new Dictionary<string, string>();
                                jgysbaF2vs.Add("项目编码", xmbh);
                                jgysbaF2vs.Add("竣工备案编号", CommonUtil.HtmlDecode(jgysbaFieldNodeList[1].InnerText.Trim()));
                                jgysbaF2vs.Add("省级竣工备案编号", CommonUtil.HtmlDecode(jgysbaFieldNodeList[2].InnerText.Trim()));
                                jgysbaF2vs.Add("实际造价（万元）", CommonUtil.HtmlDecode(jgysbaFieldNodeList[3].InnerText.Trim()));
                                jgysbaF2vs.Add("实际面积（平方米）", CommonUtil.HtmlDecode(jgysbaFieldNodeList[4].InnerText.Trim()));
                                jgysbaF2vs.Add("实际开工日期", CommonUtil.HtmlDecode(jgysbaFieldNodeList[5].InnerText.Trim()));
                                jgysbaF2vs.Add("实际竣工验收日期", CommonUtil.HtmlDecode(jgysbaFieldNodeList[6].InnerText.Trim()));
                                jgysbaCW.AddRow(jgysbaF2vs);
                            }
                        }
                        #endregion
                    }
                }
                catch (Exception ex)
                {
                    //throw ex;
                    string dir = this.RunPage.GetDetailSourceFileDir();
                    string toDir =Path.Combine(Path.GetDirectoryName(dir), "deleted");
                    string fileUrl = this.RunPage.GetFilePath(detailPageUrl, dir);
                    string toFileUrl = this.RunPage.GetFilePath(detailPageUrl, toDir);
                    File.Move(fileUrl, toFileUrl);
                    this.RunPage.InvokeAppendLogText("文件不完整，删除", LogLevelType.Error, true);
                }
            }

            mainCW.SaveToDisk();
            ztbCW.SaveToDisk();
            sgtscCW.SaveToDisk();
            htbaCW.SaveToDisk();
            sgxkCW.SaveToDisk();
            jgysbaCW.SaveToDisk();
            return true;
        }
    }
}