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
using NetDataAccess.Base.DataTransform.Address;

namespace NetDataAccess.Extended.Jianzhu.QGJZSCJGGGFWPT
{
    public class GetCompanyAddressParts : ExternalRunWebPage
    {
        private AddressTransform _AddTrans = null;
        public override bool BeforeAllGrab()
        {
            string[] parameters = this.Parameters.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            string xzqhFilePath = parameters[0];

            AddressTransform at = new AddressTransform();
            at.InitDeaultXZQHMap(xzqhFilePath);
            this._AddTrans = at;
            return base.BeforeAllGrab();
        }

        public override void GetDataByOtherAcessType(Dictionary<string, string> listRow)
        {
            try
            {
                string sourceDir = this.RunPage.GetDetailSourceFileDir();
                string detailUrl = listRow[SysConfig.DetailPageUrlFieldName];
                string detailFilePath = this.RunPage.GetFilePath(detailUrl, sourceDir);
                string address = listRow["企业经营地址"];
                string areaFullName = listRow["企业注册属地"];
                string companyName = listRow["企业名称"].Trim();

                if (!File.Exists(detailFilePath))
                {
                    string parentAreaCode = "";
                    List<string> areaCodes = this._AddTrans.GetAreaParts(areaFullName);
                    parentAreaCode = areaCodes[areaCodes.Count - 1];

                    List<string> companyparts = this._AddTrans.GetAddressParts(companyName, false);
                    if (companyparts != null && companyparts.Count > 1)
                    {
                        parentAreaCode = companyparts[0];
                    }

                    if (areaCodes != null)
                    {
                        List<string> parts = this._AddTrans.GetAddressParts(parentAreaCode, address, true);
                        if (parts == null)
                        {
                            string errorInfo = "警告: 不确定地址是否正确 . address = " + address + ", areaFullName = " + areaFullName;
                            throw new UnknownAddressException(errorInfo);
                        }
                        else
                        {
                            if (parts.Count == 1)
                            {
                                while (parentAreaCode != null && parentAreaCode.Length > 0)
                                {
                                    XZQHArea parentArea = this._AddTrans.DefaultXZQHMap.GetArea(parentAreaCode);
                                    parts.Insert(0, "code:" + parentArea.Code + ",name:" + parentArea.Name);
                                    parentAreaCode = parentArea.ParentAreaCode;
                                }
                            }

                            string addressPartStr = CommonUtil.StringArrayToString(parts.ToArray(), ";");
                            CommonUtil.CreateFileDirectory(detailFilePath);
                            FileHelper.SaveTextToFile(addressPartStr, detailFilePath);
                        }
                    }
                    else
                    {
                        throw new Exception("无法识别的地区. areaFullName = " + areaFullName);
                    }

                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public override bool AfterGrabOneCatchException(string pageUrl, Dictionary<string, string> listRow, Exception ex)
        {
            if (ex is UnknownAddressException)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private string GetAddresParts(Dictionary<string, string> listRow)
        {
            string sourceDir = this.RunPage.GetDetailSourceFileDir();
            string detailUrl = listRow[SysConfig.DetailPageUrlFieldName];
            string detailFilePath = this.RunPage.GetFilePath(detailUrl, sourceDir);
            string addressPartStr = FileHelper.GetTextFromFile(detailFilePath); 
            return addressPartStr;
        }

        public override bool AfterAllGrab(IListSheet listSheet)
        {

            String exportDir = this.RunPage.GetExportDir();

            Dictionary<string, int> resultColumnDic = new Dictionary<string, int>();
            resultColumnDic.Add("detailPageUrl", 0);
            resultColumnDic.Add("detailPageName", 1);
            resultColumnDic.Add("cookie", 2);
            resultColumnDic.Add("grabStatus", 3);
            resultColumnDic.Add("giveUpGrab", 4);
            resultColumnDic.Add("CompanyId", 5);
            resultColumnDic.Add("企业名称", 6);
            resultColumnDic.Add("统一社会信用代码", 7);
            resultColumnDic.Add("企业法定代表人", 8);
            resultColumnDic.Add("企业登记注册类型", 9);
            resultColumnDic.Add("企业注册属地", 10);
            resultColumnDic.Add("企业经营地址", 11);
            resultColumnDic.Add("addressParts", 12);  
            string resultFilePath = Path.Combine(exportDir, "企业数据_企业工商信息列表页.xlsx");
            ExcelWriter resultEW = new ExcelWriter(resultFilePath, "List", resultColumnDic, null);

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;
            Dictionary<string, string> companyDic = new Dictionary<string, string>();

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> row = listSheet.GetRow(i);
                bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
                if (!giveUp)
                {
                    string companyName = row["企业名称"].Trim().Replace("造价企业", "").Replace("测试企业", "");

                    if (!companyDic.ContainsKey(companyName))
                    {
                        Dictionary<string, string> f2vs = new Dictionary<string, string>();
                        companyDic.Add(companyName, null);

                        f2vs.Add("detailPageUrl", "https://www.tianyancha.com/search?key=" + companyName);
                        f2vs.Add("detailPageName", row["CompanyId"]);
                        f2vs.Add("CompanyId", row["CompanyId"]);
                        f2vs.Add("企业名称", companyName);
                        f2vs.Add("统一社会信用代码", row["统一社会信用代码"]);
                        f2vs.Add("企业法定代表人", row["企业法定代表人"]);
                        f2vs.Add("企业登记注册类型", row["企业登记注册类型"]);
                        f2vs.Add("企业注册属地", row["企业注册属地"]);
                        f2vs.Add("企业经营地址", row["企业经营地址"]);

                        string addressParts = this.GetAddresParts(row);
                        f2vs.Add("addressParts", addressParts);

                        resultEW.AddRow(f2vs);
                    }
                }
            }

            resultEW.SaveToDisk();

            return true;
        }
    }
}