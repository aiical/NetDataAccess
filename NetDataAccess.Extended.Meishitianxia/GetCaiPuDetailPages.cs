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
using NetDataAccess.Base.Reader;

namespace NetDataAccess.Extended.Meishitianxia
{
    public class GetCaiPuDetailPages : ExternalRunWebPage
    {
        public override bool AfterGrabOneCatchException(string pageUrl, Dictionary<string, string> listRow, Exception ex)
        {
            return true;
        }

        public override bool AfterAllGrab(IListSheet listSheet)
        {
            this.GetDetailInfos(listSheet);
            this.GetDetailMaterialInfos(listSheet);
            return true;
        }
         

        private ExcelWriter CreateDetailWriter(string subCategoryFilePath)
        {
            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>(); 
            resultColumnDic.Add("url", 0);
            resultColumnDic.Add("标题", 1);
            resultColumnDic.Add("菜谱Id", 2);
            resultColumnDic.Add("菜名", 3);
            resultColumnDic.Add("口味", 4);
            resultColumnDic.Add("工艺", 5);
            resultColumnDic.Add("耗时", 6);
            resultColumnDic.Add("难度", 7);
            resultColumnDic.Add("步骤", 8);
            ExcelWriter resultEW = new ExcelWriter(subCategoryFilePath, "List", resultColumnDic, null);
            return resultEW;
        }

        private void GetDetailInfos(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            string resultFilePath = Path.Combine(exportDir, "美食天下_菜谱详情.xlsx");

            ExcelWriter resultEW = this.CreateDetailWriter(resultFilePath);

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailPageUrl = row[SysConfig.DetailPageUrlFieldName];
                string url = row["url"];
                string name = row["name"];
                string kouWei = "";
                string gongYi = "";
                string haoShi = "";
                string nanDu = "";
                StringBuilder buZhou = new StringBuilder();
                
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    try
                    {
                        HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                        if (!htmlDoc.DocumentNode.InnerText.Contains("这个页面被人偷吃了"))
                        {
                            HtmlNode recipe_idNode = htmlDoc.DocumentNode.SelectSingleNode("//input[@id=\"recipe_id\"]");

                            if (recipe_idNode == null)
                            {
                                this.RunPage.InvokeAppendLogText("找不到recipe_idNode, 删除文件 url = " + url, LogLevelType.System, true);
                                string filePath = this.RunPage.GetFilePath(url, pageSourceDir);
                                File.Delete(filePath);

                            }
                            else
                            {

                                HtmlNode recipe_titleNode = htmlDoc.DocumentNode.SelectSingleNode("//input[@id=\"recipe_title\"]");
                                string recipe_id = CommonUtil.HtmlDecode(recipe_idNode.InnerText).Trim();
                                string recipe_title = CommonUtil.HtmlDecode(recipe_titleNode.InnerText).Trim();

                                HtmlNodeCollection cateNodes = htmlDoc.DocumentNode.SelectNodes("//div[@class=\"recipeCategory_sub_R mt30 clear\"]/ul/li");

                                foreach (HtmlNode cateNode in cateNodes)
                                {
                                    HtmlNodeCollection spanNodes = cateNode.SelectNodes("./span");
                                    string k = CommonUtil.HtmlDecode(spanNodes[1].InnerText).Trim();
                                    string v = CommonUtil.HtmlDecode(spanNodes[0].InnerText).Trim();
                                    switch (k)
                                    {
                                        //口味
                                        case "口味":
                                            kouWei = v;
                                            break;

                                        //工艺
                                        case "工艺":
                                            gongYi = v;
                                            break;

                                        //耗时
                                        case "耗时":
                                            haoShi = v;
                                            break;

                                        //难度
                                        case "难度":
                                            nanDu = v;
                                            break;
                                    }
                                }

                                //步骤
                                HtmlNodeCollection stepNodes = htmlDoc.DocumentNode.SelectNodes("//div[@class=\"recipeStep\"]/ul/li");
                                if (stepNodes != null)
                                {
                                    foreach (HtmlNode stepNode in stepNodes)
                                    {
                                        string stepText = CommonUtil.HtmlDecode(stepNode.InnerText).Trim();
                                        buZhou.Append(stepText);
                                    }
                                }


                                Dictionary<string, string> resultRow = new Dictionary<string, string>();
                                resultRow.Add("url", url);
                                resultRow.Add("标题", name);
                                resultRow.Add("菜谱Id", recipe_id);
                                resultRow.Add("菜名", recipe_title);
                                resultRow.Add("口味", kouWei);
                                resultRow.Add("工艺", gongYi);
                                resultRow.Add("耗时", haoShi);
                                resultRow.Add("难度", nanDu);
                                resultRow.Add("步骤", buZhou.ToString());
                                resultEW.AddRow(resultRow);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
            }
            resultEW.SaveToDisk();
        }



        private ExcelWriter CreateDetailMaterialWriter(string subCategoryFilePath)
        {
            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("url", 0);
            resultColumnDic.Add("标题", 1);
            resultColumnDic.Add("菜谱Id", 2);
            resultColumnDic.Add("菜名", 3);
            resultColumnDic.Add("材料类型", 4);
            resultColumnDic.Add("材料名称", 5);
            resultColumnDic.Add("用量", 6);
            ExcelWriter resultEW = new ExcelWriter(subCategoryFilePath, "List", resultColumnDic, null);
            return resultEW;
        }

        //获取主料，辅料等
        private void GetDetailMaterialInfos(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();
            string resultFilePath = Path.Combine(exportDir, "美食天下_菜谱详情_材料.xlsx");

            ExcelWriter resultEW = this.CreateDetailMaterialWriter(resultFilePath);

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailPageUrl = row[SysConfig.DetailPageUrlFieldName];
                string url = row["url"];
                string name = row["name"];

                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    try
                    {
                        HtmlAgilityPack.HtmlDocument htmlDoc = this.RunPage.GetLocalHtmlDocument(listSheet, i);
                        HtmlNode recipe_idNode = htmlDoc.DocumentNode.SelectSingleNode("//input[@id=\"recipe_id\"]");
                        HtmlNode recipe_titleNode = htmlDoc.DocumentNode.SelectSingleNode("//input[@id=\"recipe_title\"]");
                        string recipe_id = CommonUtil.HtmlDecode(recipe_idNode.InnerText).Trim();
                        string recipe_title = CommonUtil.HtmlDecode(recipe_titleNode.InnerText).Trim();

                        HtmlNodeCollection particularNodes = htmlDoc.DocumentNode.SelectNodes("//fieldset[@class=\"particulars\"]");

                        foreach (HtmlNode particularNode in particularNodes)
                        {
                            string materialType = CommonUtil.HtmlDecode(particularNode.SelectSingleNode("./legend").InnerText).Trim();
                            HtmlNodeCollection materialNodes = particularNode.SelectNodes("./div/ul/li");
                            if (materialNodes != null)
                            {
                                foreach (HtmlNode materialNode in materialNodes)
                                {
                                    HtmlNodeCollection mInfoNodes = materialNode.SelectNodes("./span");
                                    string materialName = CommonUtil.HtmlDecode(materialNodes[0].InnerText).Trim();
                                    string materialAmount = CommonUtil.HtmlDecode(materialNodes[1].InnerText).Trim();

                                    Dictionary<string, string> resultRow = new Dictionary<string, string>();
                                    resultRow.Add("url", url);
                                    resultRow.Add("标题", name);
                                    resultRow.Add("菜谱Id", recipe_id);
                                    resultRow.Add("菜名", recipe_title);
                                    resultRow.Add("材料类型", materialType);
                                    resultRow.Add("材料名称", materialName);
                                    resultRow.Add("用量", materialAmount);
                                    resultEW.AddRow(resultRow);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        this.RunPage.InvokeAppendLogText(ex.Message + "url = " + url + ", name = " + name, LogLevelType.Error, true);
                        throw ex;
                    }
                }
            }
            resultEW.SaveToDisk();
        } 
    }
}