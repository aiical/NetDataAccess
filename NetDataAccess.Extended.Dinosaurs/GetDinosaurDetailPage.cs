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
using System.Security.Cryptography;

namespace NetDataAccess.Extended.Dinosaurs
{
    public class GetDinosaurDetailPage : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            this.GetDetailPage(listSheet);
            this.GetImagePageUrls(listSheet);
            this.GetTranslatePageUrls(listSheet);
            return true;
        }

        private void GetTranslatePageUrls(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("name", 5);
            resultColumnDic.Add("length", 6);
            resultColumnDic.Add("description", 7);
            resultColumnDic.Add("diet", 8);
            resultColumnDic.Add("country", 9);
            resultColumnDic.Add("period", 10);
            resultColumnDic.Add("teeth", 11);
            resultColumnDic.Add("how_it_moved", 12);
            resultColumnDic.Add("food", 13);
            resultColumnDic.Add("content", 14);
            resultColumnDic.Add("taxonomy", 15);
            resultColumnDic.Add("type_species", 16);
            resultColumnDic.Add("url", 17);
            string resultFilePath = Path.Combine(exportDir, "www.nhm.ac.uk恐龙中文信息.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            Dictionary<string, string> urlDic = new Dictionary<string, string>();
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"];
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string localFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir);

                    try
                    {
                        HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                        htmlDoc.Load(localFilePath, Encoding.GetEncoding("utf-8"));
                        string name = row["name"];
                        string length = "";
                        string description = "";
                        string diet = "";
                        string country = "";
                        string period = "";
                        string teeth = "";
                        string how_it_moved = "";
                        string food = "";
                        string content = "";
                        string taxonomy = "";
                        string type_species = "";

                        HtmlNode descriptionNode = htmlDoc.DocumentNode.SelectSingleNode("//div[@class='dinosaur--description']");
                        length = descriptionNode.GetAttributeValue("data-dino-length", "");
                        description = CommonUtil.HtmlDecode(descriptionNode.InnerText.Trim());

                        HtmlNode dietNode = htmlDoc.DocumentNode.SelectSingleNode("//div[@class='dinosaur--diet']");
                        if (dietNode != null)
                        {
                            diet = CommonUtil.HtmlDecode(dietNode.InnerText.Trim()).Replace("Diet:", "").Trim();
                        }

                        HtmlNode countryNode = htmlDoc.DocumentNode.SelectSingleNode("//div[@class='dinosaur--country']");
                        if (countryNode != null)
                        {
                            country = CommonUtil.HtmlDecode(countryNode.InnerText.Trim()).Replace("Country:", "").Trim();
                        }

                        HtmlNode periodNode = htmlDoc.DocumentNode.SelectSingleNode("//div[@class='dinosaur--period']");
                        if (periodNode != null)
                        {
                            period = CommonUtil.HtmlDecode(periodNode.InnerText.Trim()).Replace("Period:", "").Trim();
                        }

                        HtmlNodeCollection infoPNodes = htmlDoc.DocumentNode.SelectNodes("//div[starts-with(@class, 'dinosaur--info-container')]/p");
                        if (infoPNodes != null)
                        {
                            for (int j = 0; j < infoPNodes.Count; j++)
                            {
                                HtmlNode infoPNode = infoPNodes[j];
                                string infoPText = CommonUtil.HtmlDecode(infoPNode.InnerText.Trim()).Trim();
                                if (infoPText.StartsWith("Teeth:"))
                                {
                                    teeth = infoPText.Replace("Teeth:", "").Trim();
                                }
                                else if (infoPText.StartsWith("How it moved:"))
                                {
                                    how_it_moved = infoPText.Replace("How it moved:", "").Trim();
                                }
                                else if (infoPText.StartsWith("Food:"))
                                {
                                    food = infoPText.Replace("Food:", "").Trim();
                                }
                            }
                        }

                        HtmlNode contentNode = htmlDoc.DocumentNode.SelectSingleNode("//div[starts-with(@class, 'dinosaur--content-container')]/p");
                        if (contentNode != null)
                        {
                            content = CommonUtil.HtmlDecode(contentNode.InnerText.Trim()).Trim();
                        }

                        HtmlNodeCollection taxonomyNodes = htmlDoc.DocumentNode.SelectNodes("//div[@class='dinosaur--taxonomy-detail']");
                        if (taxonomyNodes != null)
                        {
                            for (int j = 0; j < taxonomyNodes.Count; j++)
                            {
                                HtmlNode taxonomyNode = taxonomyNodes[j];
                                string taxonomyText = CommonUtil.HtmlDecode(taxonomyNode.InnerText.Trim()).Trim();
                                if (taxonomyText.StartsWith("Taxonomy:"))
                                {
                                    taxonomy = taxonomyText.Replace("Taxonomy:", "").Trim();
                                }
                                else if (taxonomyText.StartsWith("Type species:"))
                                {
                                    type_species = taxonomyText.Replace("Type species:", "").Trim();
                                }
                            }
                        }
                        string appId = "20180507000154860";
                        string salt = DateTime.Now.ToString("yyyyMMddHHmmssfff");
                        string key = "hYitn9aXvMRVQfMM2xMX";
                        MD5 md5 = new MD5CryptoServiceProvider();
                        byte[] result = Encoding.GetEncoding("utf-8").GetBytes(appId + name + salt + key);
                        byte[] output = md5.ComputeHash(result);
                        string sign = BitConverter.ToString(output).Replace("-", "").ToLower();
                        string trUrl = "http://api.fanyi.baidu.com/api/trans/vip/translate?q=" + name + "&from=en&to=zh&appid=" + appId + "&salt=" + salt + "&sign=" + sign;

                        Dictionary<string, string> f2vs = new Dictionary<string, string>();
                        f2vs.Add("detailPageUrl", trUrl);
                        f2vs.Add("detailPageName", trUrl);
                        f2vs.Add("name", name);
                        f2vs.Add("length", length);
                        f2vs.Add("description", description);
                        f2vs.Add("diet", diet);
                        f2vs.Add("country", country);
                        f2vs.Add("period", period);
                        f2vs.Add("teeth", teeth);
                        f2vs.Add("how_it_moved", how_it_moved);
                        f2vs.Add("food", food);
                        f2vs.Add("content", content);
                        f2vs.Add("taxonomy", taxonomy);
                        f2vs.Add("type_species", type_species);
                        f2vs.Add("url", detailUrl);
                        resultEW.AddRow(f2vs);
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
            }
            resultEW.SaveToDisk();
        }


        private void GetDetailPage(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("name", 0);
            resultColumnDic.Add("length", 1);
            resultColumnDic.Add("description", 2);
            resultColumnDic.Add("diet", 3);
            resultColumnDic.Add("country", 4);
            resultColumnDic.Add("period", 5);
            resultColumnDic.Add("teeth", 6);
            resultColumnDic.Add("how_it_moved", 7);
            resultColumnDic.Add("food", 8);
            resultColumnDic.Add("content", 9);
            resultColumnDic.Add("taxonomy", 10);
            resultColumnDic.Add("type_species", 11);
            resultColumnDic.Add("url", 12);
            string resultFilePath = Path.Combine(exportDir, "www.nhm.ac.uk恐龙信息.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            Dictionary<string, string> urlDic = new Dictionary<string, string>();
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"];
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string localFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir);

                    try
                    {
                        HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                        htmlDoc.Load(localFilePath, Encoding.GetEncoding("utf-8"));
                        string name = row["name"];
                        string length = "";
                        string description = "";
                        string diet = "";
                        string country = "";
                        string period = "";
                        string teeth = "";
                        string how_it_moved = "";
                        string food = "";
                        string content = "";
                        string taxonomy = "";
                        string type_species = "";

                        HtmlNode descriptionNode = htmlDoc.DocumentNode.SelectSingleNode("//div[@class='dinosaur--description']");
                        length = descriptionNode.GetAttributeValue("data-dino-length", "");
                        description = CommonUtil.HtmlDecode(descriptionNode.InnerText.Trim());

                        HtmlNode dietNode = htmlDoc.DocumentNode.SelectSingleNode("//div[@class='dinosaur--diet']");
                        if (dietNode != null)
                        {
                            diet = CommonUtil.HtmlDecode(dietNode.InnerText.Trim()).Replace("Diet:", "").Trim();
                        }

                        HtmlNode countryNode = htmlDoc.DocumentNode.SelectSingleNode("//div[@class='dinosaur--country']");
                        if (countryNode != null)
                        {
                            country = CommonUtil.HtmlDecode(countryNode.InnerText.Trim()).Replace("Country:", "").Trim();
                        }

                        HtmlNode periodNode = htmlDoc.DocumentNode.SelectSingleNode("//div[@class='dinosaur--period']");
                        if (periodNode != null)
                        {
                            period = CommonUtil.HtmlDecode(periodNode.InnerText.Trim()).Replace("Period:", "").Trim();
                        }

                        HtmlNodeCollection infoPNodes = htmlDoc.DocumentNode.SelectNodes("//div[starts-with(@class, 'dinosaur--info-container')]/p");
                        if (infoPNodes != null)
                        {
                            for (int j = 0; j < infoPNodes.Count; j++)
                            {
                                HtmlNode infoPNode = infoPNodes[j];
                                string infoPText = CommonUtil.HtmlDecode(infoPNode.InnerText.Trim()).Trim();
                                if (infoPText.StartsWith("Teeth:"))
                                {
                                    teeth = infoPText.Replace("Teeth:", "").Trim();
                                }
                                else if (infoPText.StartsWith("How it moved:"))
                                {
                                    how_it_moved = infoPText.Replace("How it moved:", "").Trim();
                                }
                                else if (infoPText.StartsWith("Food:"))
                                {
                                    food = infoPText.Replace("Food:", "").Trim();
                                }
                            }
                        }

                        HtmlNode contentNode = htmlDoc.DocumentNode.SelectSingleNode("//div[starts-with(@class, 'dinosaur--content-container')]/p");
                        if (contentNode != null)
                        {
                            content = CommonUtil.HtmlDecode(contentNode.InnerText.Trim()).Trim();
                        }

                        HtmlNodeCollection taxonomyNodes = htmlDoc.DocumentNode.SelectNodes("//div[@class='dinosaur--taxonomy-detail']");
                        if (taxonomyNodes != null)
                        {
                            for (int j = 0; j < taxonomyNodes.Count; j++)
                            {
                                HtmlNode taxonomyNode = taxonomyNodes[j];
                                string taxonomyText = CommonUtil.HtmlDecode(taxonomyNode.InnerText.Trim()).Trim();
                                if (taxonomyText.StartsWith("Taxonomy:"))
                                {
                                    taxonomy = taxonomyText.Replace("Taxonomy:", "").Trim();
                                }
                                else if (taxonomyText.StartsWith("Type species:"))
                                {
                                    type_species = taxonomyText.Replace("Type species:", "").Trim();
                                }
                            }
                        }

                        Dictionary<string, string> f2vs = new Dictionary<string, string>();
                        f2vs.Add("name", name);
                        f2vs.Add("length", length);
                        f2vs.Add("description", description);
                        f2vs.Add("diet", diet);
                        f2vs.Add("country", country);
                        f2vs.Add("period", period);
                        f2vs.Add("teeth", teeth);
                        f2vs.Add("how_it_moved", how_it_moved);
                        f2vs.Add("food", food);
                        f2vs.Add("content", content);
                        f2vs.Add("taxonomy", taxonomy);
                        f2vs.Add("type_species", type_species);
                        f2vs.Add("url", detailUrl);
                        resultEW.AddRow(f2vs);
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
            }
            resultEW.SaveToDisk();
        }


        private void GetImagePageUrls(IListSheet listSheet)
        {
            String exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("name", 5);
            string resultFilePath = Path.Combine(exportDir, "www.nhm.ac.uk恐龙图片.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);
            Dictionary<string, string> urlDic = new Dictionary<string, string>();
            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                string detailUrl = row["detailPageUrl"];
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string localFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir);

                    try
                    {
                        HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                        htmlDoc.Load(localFilePath, Encoding.GetEncoding("utf-8"));
                        string name = row["name"];

                        HtmlNode imgNode = htmlDoc.DocumentNode.SelectSingleNode("//img[@class='dinosaur--image']");
                        string imageUrl = imgNode.GetAttributeValue("src", ""); 

                        Dictionary<string, string> f2vs = new Dictionary<string, string>();
                        f2vs.Add("detailPageUrl", imageUrl);
                        f2vs.Add("detailPageName", imageUrl);
                        f2vs.Add("name", name); 
                        resultEW.AddRow(f2vs);
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
            }
            resultEW.SaveToDisk();
        }
         
    }
}