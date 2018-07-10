using NetDataAccess.Base.EnumTypes;
using System;
using System.Collections.Generic;
using System.Text;

namespace NetDataAccess.Base.Definition
{
    /// <summary>
    /// 登录页
    /// </summary>
    public class Proj_LoginPageInfo : IProj_XmlConfig
    {
        #region LoginUrl
        private string _LoginUrl = "";
        /// <summary>
        /// LoginUrl
        /// </summary>
        public string LoginUrl
        {
            set
            {
                _LoginUrl = value;
            }
            get
            {
                return _LoginUrl;
            }
        }
        #endregion

        #region LoginName
        private string _LoginName = "";
        /// <summary>
        /// LoginName
        /// </summary>
        public string LoginName
        {
            set
            {
                _LoginName = value;
            }
            get
            {
                return _LoginName;
            }
        }
        #endregion

        #region LoginPwdValue
        private string _LoginPwdValue = "";
        /// <summary>
        /// LoginPwdValue
        /// </summary>
        public string LoginPwdValue
        {
            set
            {
                _LoginPwdValue = value;
            }
            get
            {
                return _LoginPwdValue;
            }
        }
        #endregion

        #region LoginNameCtrlPath
        private string _LoginNameCtrlPath = "";
        /// <summary>
        /// LoginNameCtrlPath
        /// </summary>
        public string LoginNameCtrlPath
        {
            set
            {
                _LoginNameCtrlPath = value;
            }
            get
            {
                return _LoginNameCtrlPath;
            }
        }
        #endregion

        #region LoginPwdCtrlPath
        private string _LoginPwdCtrlPath = "";
        /// <summary>
        /// LoginPwdCtrlPath
        /// </summary>
        public string LoginPwdCtrlPath
        {
            set
            {
                _LoginPwdCtrlPath = value;
            }
            get
            {
                return _LoginPwdCtrlPath;
            }
        }
        #endregion

        #region LoginBtnPath
        private string _LoginBtnPath = "";
        /// <summary>
        /// LoginBtnPath
        /// </summary>
        public string LoginBtnPath
        {
            set
            {
                _LoginBtnPath = value;
            }
            get
            {
                return _LoginBtnPath;
            }
        }
        #endregion

        #region 数据获取方式
        private Proj_DataAccessType _DataAccessType = Proj_DataAccessType.WebBrowserHtml;
        /// <summary>
        /// 数据获取方式
        /// </summary>
        public Proj_DataAccessType DataAccessType
        {
            set
            {
                _DataAccessType = value;
            }
            get
            {
                return _DataAccessType;
            }
        }
        #endregion

        #region 是否使用代理
        private bool _NeedProxy = false;
        /// <summary>
        /// 是否使用代理
        /// </summary>
        public bool NeedProxy
        {
            set
            {
                _NeedProxy = value;
            }
            get
            {
                return _NeedProxy;
            }
        }
        #endregion

        #region 自定放弃使用返回有问题的代理
        private bool _AutoAbandonDisableProxy = false;
        /// <summary>
        /// 自定放弃使用返回有问题的代理
        /// </summary>
        public bool AutoAbandonDisableProxy
        {
            set
            {
                _AutoAbandonDisableProxy = value;
            }
            get
            {
                return _AutoAbandonDisableProxy;
            }
        }
        #endregion 
         
    }
}
