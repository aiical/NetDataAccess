using System;
using System.Collections.Generic;
using System.Text;
using NetDataAccess.Base.EnumTypes;
using System.Xml;
using NetDataAccess.Base.Common;
using NetDataAccess.Base.Config;

namespace NetDataAccess.Base.Definition
{
    /// <summary>
    /// 项目
    /// </summary>
    public class Proj_Main
    {
        #region Id
        private string _Id;
        public string Id
        {
            get
            {
                return this._Id;
            }
            set
            {
                this._Id = value;
            }
        }
        #endregion 

        #region Description
        private string _Description;
        public string Description
        {
            get
            {
                return this._Description;
            }
            set
            {
                this._Description = value;
            }
        }
        #endregion 

        #region Name
        private string _Name;
        public string Name
        {
            get
            {
                return this._Name;
            }
            set
            {
                this._Name = value;
            }
        }
        #endregion 

        #region Group_Id
        private string _Group_Id;
        public string Group_Id
        {
            get
            {
                return this._Group_Id;
            }
            set
            {
                this._Group_Id = value;
            }
        }
        #endregion  

        #region LoginType
        private LoginLevelType _LoginType;
        public LoginLevelType LoginType
        {
            get
            {
                return this._LoginType;
            }
            set
            {
                this._LoginType = value;
            }
        }
        #endregion

        #region LoginPageInfo
        private string _LoginPageInfo;
        public string LoginPageInfo
        {
            get
            {
                return this._LoginPageInfo;
            }
            set
            {
                this._LoginPageInfo = value;
            }
        }

        private IProj_XmlConfig _LoginPageInfoObject;
        public IProj_XmlConfig LoginPageInfoObject
        {
            get
            {
                return this._LoginPageInfoObject;
            }
            set
            {
                this._LoginPageInfoObject = value;
            }
        }
        #endregion

        #region DetailGrabType
        private DetailGrabType _DetailGrabType;
        public DetailGrabType DetailGrabType
        {
            get
            {
                return this._DetailGrabType;
            }
            set
            {
                this._DetailGrabType = value;
            }
        }
        #endregion

        #region DetailGrabInfo
        private string _DetailGrabInfo;
        public string DetailGrabInfo
        {
            get
            {
                return this._DetailGrabInfo;
            }
            set
            {
                this._DetailGrabInfo = value;
            }
        }

        private IProj_FieldConfig _DetailGrabInfoObject;
        public IProj_FieldConfig DetailGrabInfoObject
        {
            get
            {
                return this._DetailGrabInfoObject;
            }
            set
            {
                this._DetailGrabInfoObject = value;
            }
        }
        #endregion

        #region ProgramAfterGrabAll
        private string _ProgramAfterGrabAll;
        public string ProgramAfterGrabAll
        {
            get
            {
                return this._ProgramAfterGrabAll;
            }
            set
            {
                this._ProgramAfterGrabAll = value;
            }
        }

        private IProj_XmlConfig _ProgramAfterGrabAllObject;
        public IProj_XmlConfig ProgramAfterGrabAllObject
        {
            get
            {
                return this._ProgramAfterGrabAllObject;
            }
            set
            {
                this._ProgramAfterGrabAllObject = value;
            }
        }
        #endregion

        #region ProgramExternalRun
        private string _ProgramExternalRun;
        public string ProgramExternalRun
        {
            get
            {
                return this._ProgramExternalRun;
            }
            set
            {
                this._ProgramExternalRun = value;
            }
        }

        private IProj_XmlConfig _ProgramExternalRunObject;
        public IProj_XmlConfig ProgramExternalRunObject
        {
            get
            {
                return this._ProgramExternalRunObject;
            }
            set
            {
                this._ProgramExternalRunObject = value;
            }
        }
        #endregion

        #region 格式化设置，在Task运行前执行
        public bool Format()
        {
            try
            { 
                if(!CommonUtil.IsNullOrBlank( this.LoginPageInfo))
                {
                    Proj_LoginPageInfo loginObj = new Proj_LoginPageInfo();
                    XmlDocument xmlDoc = new XmlDocument();
                    xmlDoc.LoadXml(this.LoginPageInfo);
                    XmlElement rootElement = xmlDoc.DocumentElement;
                    loginObj.LoginUrl = rootElement.Attributes["LoginUrl"].Value;
                    loginObj.LoginBtnPath = rootElement.Attributes["LoginBtnPath"].Value;
                    loginObj.LoginName = rootElement.Attributes["LoginName"].Value;
                    loginObj.LoginNameCtrlPath = rootElement.Attributes["LoginNameCtrlPath"].Value;
                    loginObj.LoginPwdCtrlPath = rootElement.Attributes["LoginPwdCtrlPath"].Value;
                    loginObj.LoginPwdValue = rootElement.Attributes["LoginPwdValue"].Value;
                    loginObj.DataAccessType = rootElement.Attributes["DataAccessType"] == null ? loginObj.DataAccessType : (Proj_DataAccessType)Enum.Parse(typeof(Proj_DataAccessType), rootElement.Attributes["DataAccessType"].Value);
                    loginObj.NeedProxy = rootElement.Attributes["NeedProxy"] == null ? false : bool.Parse(rootElement.Attributes["NeedProxy"].Value);
                    loginObj.AutoAbandonDisableProxy = rootElement.Attributes["AutoAbandonDisableProxy"] == null ? true : bool.Parse(rootElement.Attributes["AutoAbandonDisableProxy"].Value);
                    this.LoginPageInfoObject = loginObj;
                } 

                if (!CommonUtil.IsNullOrBlank(this.DetailGrabInfo))
                {
                    switch (this.DetailGrabType)
                    {
                        case DetailGrabType.SingleLineType:
                        case DetailGrabType.ProgramType:
                            {
                                Proj_Detail_SingleLine detailObj = new Proj_Detail_SingleLine();
                                XmlDocument xmlDoc = new XmlDocument();
                                xmlDoc.LoadXml(this.DetailGrabInfo);
                                XmlElement rootElement = xmlDoc.DocumentElement;
                                detailObj.IntervalAfterLoaded = decimal.Parse(rootElement.Attributes["IntervalAfterLoaded"].Value);
                                detailObj.DataAccessType = rootElement.Attributes["DataAccessType"] == null ? detailObj.DataAccessType : (Proj_DataAccessType)Enum.Parse(typeof(Proj_DataAccessType), rootElement.Attributes["DataAccessType"].Value);
                                detailObj.NeedProxy = rootElement.Attributes["NeedProxy"] == null ? false : bool.Parse(rootElement.Attributes["NeedProxy"].Value);
                                detailObj.AutoAbandonDisableProxy = rootElement.Attributes["AutoAbandonDisableProxy"] == null ? true : bool.Parse(rootElement.Attributes["AutoAbandonDisableProxy"].Value);
                                detailObj.IntervalDetailPageSave = rootElement.Attributes["IntervalDetailPageSave"] == null ? SysConfig.IntervalDetailPageSave : int.Parse(rootElement.Attributes["IntervalDetailPageSave"].Value);
                                detailObj.StartPageIndex = rootElement.Attributes["StartPageIndex"] == null ? 0 : int.Parse(rootElement.Attributes["StartPageIndex"].Value);
                                detailObj.EndPageIndex = rootElement.Attributes["EndPageIndex"] == null ? 0 : int.Parse(rootElement.Attributes["EndPageIndex"].Value);
                                detailObj.SaveFileDirectory = rootElement.Attributes["SaveFileDirectory"] == null ? "" : rootElement.Attributes["SaveFileDirectory"].Value;
                                detailObj.ExportType = rootElement.Attributes["ExportType"] == null ? ExportType.Excel : (ExportType)Enum.Parse(typeof(ExportType), rootElement.Attributes["ExportType"].Value);
                                detailObj.AllowAutoGiveUp = rootElement.Attributes["AllowAutoGiveUp"] == null ? false : bool.Parse(rootElement.Attributes["AllowAutoGiveUp"].Value);
                                detailObj.NeedPartDir = rootElement.Attributes["NeedPartDir"] == null ? false : bool.Parse(rootElement.Attributes["NeedPartDir"].Value);
                                detailObj.ThreadCount = rootElement.Attributes["ThreadCount"] == null ? 5 : int.Parse(rootElement.Attributes["ThreadCount"].Value);
                                detailObj.RequestTimeout = rootElement.Attributes["RequestTimeout"] == null ? SysConfig.WebPageRequestTimeout : int.Parse(rootElement.Attributes["RequestTimeout"].Value);
                                detailObj.Encoding = rootElement.Attributes["Encoding"] == null ? SysConfig.WebPageEncoding : rootElement.Attributes["Encoding"].Value;
                                detailObj.XRequestedWith = rootElement.Attributes["XRequestedWith"] == null ? "" : rootElement.Attributes["XRequestedWith"].Value;
                                detailObj.IntervalProxyRequest = rootElement.Attributes["IntervalProxyRequest"] == null ? detailObj.IntervalProxyRequest : int.Parse(rootElement.Attributes["IntervalProxyRequest"].Value);
                                detailObj.BrowserType = rootElement.Attributes["BrowserType"] == null ? WebBrowserType.IE : (WebBrowserType)Enum.Parse(typeof(WebBrowserType), rootElement.Attributes["BrowserType"].Value);
                                   
                                XmlNode completeCheckListNode = rootElement.SelectSingleNode("CompleteChecks") == null ? null : rootElement.SelectSingleNode("CompleteChecks");
                                if (completeCheckListNode != null)
                                {
                                    detailObj.CompleteChecks = new Proj_CompleteCheckList();
                                    detailObj.CompleteChecks.AndCondition = completeCheckListNode.Attributes["AndCondition"] == null ? detailObj.CompleteChecks.AndCondition : bool.Parse(completeCheckListNode.Attributes["AndCondition"].Value);
                                    XmlNodeList completeCheckList = completeCheckListNode.ChildNodes;
                                    foreach (XmlNode completeCheckNode in completeCheckList)
                                    {
                                        Proj_CompleteCheck completeCheck = new Proj_CompleteCheck();
                                        completeCheck.CheckValue = completeCheckNode.Attributes["CheckValue"] == null ? "" : completeCheckNode.Attributes["CheckValue"].Value;
                                        completeCheck.CheckType = completeCheckNode.Attributes["CheckType"] == null ? DocumentCompleteCheckType.BrowserCompleteEvent : (DocumentCompleteCheckType)Enum.Parse(typeof(DocumentCompleteCheckType), completeCheckNode.Attributes["CheckType"].Value);
                                        detailObj.CompleteChecks.Add(completeCheck);
                                    }
                                }
                               
                                XmlNodeList fieldNodeList = rootElement.SelectSingleNode("Fields") == null ? null : rootElement.SelectSingleNode("Fields").ChildNodes;
                                if (fieldNodeList != null)
                                {
                                    foreach (XmlNode fieldNode in fieldNodeList)
                                    {
                                        Proj_Detail_Field field = new Proj_Detail_Field();
                                        field.Name = fieldNode.Attributes["Name"].Value;
                                        field.Path = fieldNode.Attributes["Path"] == null ? "" : fieldNode.Attributes["Path"].Value;
                                        field.AttributeName = fieldNode.Attributes["AttributeName"] == null ? "" : fieldNode.Attributes["AttributeName"].Value;
                                        field.NeedAllHtml = fieldNode.Attributes["NeedAllHtml"] == null ? false : "Y".Equals(fieldNode.Attributes["NeedAllHtml"].Value.ToUpper());
                                        field.ColumnWidth = fieldNode.Attributes["ColumnWidth"] == null ? field.ColumnWidth : int.Parse(fieldNode.Attributes["ColumnWidth"].Value);
                                        detailObj.Fields.Add(field);
                                    }
                                }
                                this.DetailGrabInfoObject = detailObj;
                            }
                            break;
                        case DetailGrabType.MultiLineType:
                            {
                                Proj_Detail_MultiLine detailObj = new Proj_Detail_MultiLine();
                                XmlDocument xmlDoc = new XmlDocument();
                                xmlDoc.LoadXml(this.DetailGrabInfo);
                                XmlElement rootElement = xmlDoc.DocumentElement;
                                detailObj.IntervalAfterLoaded = int.Parse(rootElement.Attributes["IntervalAfterLoaded"].Value);
                                detailObj.MultiCtrlPath = rootElement.Attributes["MultiCtrlPath"].Value;
                                detailObj.DataAccessType = rootElement.Attributes["DataAccessType"] == null ? detailObj.DataAccessType : (Proj_DataAccessType)Enum.Parse(typeof(Proj_DataAccessType), rootElement.Attributes["DataAccessType"].Value);
                                detailObj.NeedProxy = rootElement.Attributes["NeedProxy"] == null ? false : bool.Parse(rootElement.Attributes["NeedProxy"].Value);
                                detailObj.AutoAbandonDisableProxy = rootElement.Attributes["AutoAbandonDisableProxy"] == null ? true : bool.Parse(rootElement.Attributes["AutoAbandonDisableProxy"].Value);
                                detailObj.IntervalDetailPageSave = rootElement.Attributes["IntervalDetailPageSave"] == null ? SysConfig.IntervalDetailPageSave : int.Parse(rootElement.Attributes["IntervalDetailPageSave"].Value);
                                detailObj.StartPageIndex = rootElement.Attributes["StartPageIndex"] == null ? 0 : int.Parse(rootElement.Attributes["StartPageIndex"].Value);
                                detailObj.EndPageIndex = rootElement.Attributes["EndPageIndex"] == null ? 0 : int.Parse(rootElement.Attributes["EndPageIndex"].Value);
                                detailObj.SaveFileDirectory = rootElement.Attributes["SaveFileDirectory"] == null ? "" : rootElement.Attributes["SaveFileDirectory"].Value;
                                detailObj.ExportType = rootElement.Attributes["ExportType"] == null ? ExportType.Excel : (ExportType)Enum.Parse(typeof(ExportType), rootElement.Attributes["ExportType"].Value);
                                detailObj.AllowAutoGiveUp = rootElement.Attributes["AllowAutoGiveUp"] == null ? false : bool.Parse(rootElement.Attributes["AllowAutoGiveUp"].Value);
                                detailObj.NeedPartDir = rootElement.Attributes["NeedPartDir"] == null ? false : bool.Parse(rootElement.Attributes["NeedPartDir"].Value);
                                detailObj.ThreadCount = rootElement.Attributes["ThreadCount"] == null ? 5 : int.Parse(rootElement.Attributes["ThreadCount"].Value);
                                detailObj.RequestTimeout = rootElement.Attributes["RequestTimeout"] == null ? SysConfig.WebPageRequestTimeout : int.Parse(rootElement.Attributes["RequestTimeout"].Value);
                                detailObj.Encoding = rootElement.Attributes["Encoding"] == null ? SysConfig.WebPageEncoding : rootElement.Attributes["Encoding"].Value;
                                detailObj.XRequestedWith = rootElement.Attributes["XRequestedWith"] == null ? "" : rootElement.Attributes["XRequestedWith"].Value;
                                detailObj.IntervalProxyRequest = rootElement.Attributes["IntervalProxyRequest"] == null ? detailObj.IntervalProxyRequest : int.Parse(rootElement.Attributes["IntervalProxyRequest"].Value);
                                detailObj.BrowserType = rootElement.Attributes["BrowserType"] == null ? WebBrowserType.IE : (WebBrowserType)Enum.Parse(typeof(WebBrowserType), rootElement.Attributes["BrowserType"].Value);
                             
                                XmlNode completeCheckListNode = rootElement.SelectSingleNode("CompleteChecks") == null ? null : rootElement.SelectSingleNode("CompleteChecks");
                                if (completeCheckListNode != null)
                                {
                                    detailObj.CompleteChecks = new Proj_CompleteCheckList();
                                    detailObj.CompleteChecks.AndCondition = completeCheckListNode.Attributes["AndCondition"] == null ? detailObj.CompleteChecks.AndCondition : bool.Parse(completeCheckListNode.Attributes["AndCondition"].Value);
                                    XmlNodeList completeCheckList = completeCheckListNode.ChildNodes;
                                    foreach (XmlNode completeCheckNode in completeCheckList)
                                    {
                                        Proj_CompleteCheck completeCheck = new Proj_CompleteCheck();
                                        completeCheck.CheckValue = completeCheckNode.Attributes["CheckValue"] == null ? "" : completeCheckNode.Attributes["CheckValue"].Value;
                                        completeCheck.CheckType = completeCheckNode.Attributes["CheckType"] == null ? DocumentCompleteCheckType.BrowserCompleteEvent : (DocumentCompleteCheckType)Enum.Parse(typeof(DocumentCompleteCheckType), completeCheckNode.Attributes["CheckType"].Value);
                                        detailObj.CompleteChecks.Add(completeCheck);
                                    }
                                }

                                XmlNodeList fieldNodeList = rootElement.SelectSingleNode("Fields").ChildNodes;
                                foreach (XmlNode fieldNode in fieldNodeList)
                                {
                                    Proj_Detail_Field field = new Proj_Detail_Field();
                                    field.Name = fieldNode.Attributes["Name"].Value;
                                    field.Path = fieldNode.Attributes["Path"] == null ? field.Name : fieldNode.Attributes["Path"].Value;
                                    field.AttributeName = fieldNode.Attributes["AttributeName"] == null ? "" : fieldNode.Attributes["AttributeName"].Value;
                                    field.IsAbsolute = fieldNode.Attributes["IsAbsolute"] == null ? false : "Y".Equals(fieldNode.Attributes["IsAbsolute"].Value.ToUpper());
                                    field.NeedAllHtml = fieldNode.Attributes["NeedAllHtml"] == null ? false : "Y".Equals(fieldNode.Attributes["NeedAllHtml"].Value.ToUpper());
                                    field.ColumnWidth = fieldNode.Attributes["ColumnWidth"] == null ? field.ColumnWidth : int.Parse(fieldNode.Attributes["ColumnWidth"].Value);
                                    detailObj.Fields.Add(field);
                                }
                                this.DetailGrabInfoObject = detailObj;
                            }
                            break;
                            /*
                        case DetailGrabType.ProgramType:
                            if (!CommonUtil.IsNullOrBlank(this.DetailGrabInfo))
                            {
                                this.DetailGrabInfoObject = GetCustomProgram(this.DetailGrabInfo);
                            }
                            break;
                             * */
                    }
                }
                else
                {
                    if (this.DetailGrabType != EnumTypes.DetailGrabType.NoneDetailPage)
                    {
                        CommonUtil.Alert("错误", "没有设置详情页.");
                        return false;
                    }
                }

                if (!CommonUtil.IsNullOrBlank(this.ProgramAfterGrabAll))
                { 
                    this.ProgramAfterGrabAllObject = GetCustomProgram(this.ProgramAfterGrabAll);
                }

                if (!CommonUtil.IsNullOrBlank(this.ProgramExternalRun))
                {
                    this.ProgramExternalRunObject = GetCustomProgram(this.ProgramExternalRun);
                }
                return true;
            }
            catch (Exception ex)
            {
                CommonUtil.Alert("错误", "格式化设置失败.\r\n" + ex.Message);
                return false;
            }
        }

        private Proj_CustomProgram GetCustomProgram(string xml)
        {
            Proj_CustomProgram obj = new Proj_CustomProgram();
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(xml);
            XmlElement rootElement = xmlDoc.DocumentElement;
            obj.AssemblyName = rootElement.Attributes["AssemblyName"].Value;
            obj.NamespaceName = rootElement.Attributes["NamespaceName"].Value;
            obj.ClassName = rootElement.Attributes["ClassName"].Value;
            obj.Parameters = rootElement.Attributes["Parameters"].Value;
            obj.NeedProxy = rootElement.Attributes["NeedProxy"] == null ? false : bool.Parse(rootElement.Attributes["NeedProxy"].Value);
            obj.AutoAbandonDisableProxy = rootElement.Attributes["AutoAbandonDisableProxy"] == null ? true : bool.Parse(rootElement.Attributes["AutoAbandonDisableProxy"].Value);
            obj.SaveSourceFile = rootElement.Attributes["SaveSourceFile"] == null ? false : bool.Parse(rootElement.Attributes["SaveSourceFile"].Value);
            obj.SaveFileDirectory = rootElement.Attributes["SaveFileDirectory"] == null ? "" : rootElement.Attributes["SaveFileDirectory"].Value;
            obj.Parameters = rootElement.Attributes["Parameters"].Value;
            XmlNodeList fieldNodeList = rootElement.SelectSingleNode("Fields") == null ? null : rootElement.SelectSingleNode("Fields").ChildNodes;
            if (fieldNodeList != null)
            {
                foreach (XmlNode fieldNode in fieldNodeList)
                {
                    Proj_Detail_Field field = new Proj_Detail_Field();
                    field.Name = fieldNode.Attributes["Name"].Value;
                    field.ColumnWidth = fieldNode.Attributes["ColumnWidth"] == null ? field.ColumnWidth : int.Parse(fieldNode.Attributes["ColumnWidth"].Value);
                    obj.Fields.Add(field);
                }
            }
            return obj;
        }
        #endregion 
    }
}