using System;
using System.Collections.Generic;
using System.Text;

namespace NetDataAccess.Base.Definition
{
    /// <summary>
    /// 自定义程序入口
    /// </summary>
    public class Proj_CustomProgram : IProj_FieldConfig 
    {
        #region AssemblyName
        private string _AssemblyName;
        /// <summary>
        /// AssemblyName
        /// </summary>
        public string AssemblyName
        {
            set
            {
                _AssemblyName = value;
            }
            get
            {
                return _AssemblyName;
            }
        }
        #endregion

        #region NamespaceName
        private string _NamespaceName;
        /// <summary>
        /// NamespaceName
        /// </summary>
        public string NamespaceName
        {
            set
            {
                _NamespaceName = value;
            }
            get
            {
                return _NamespaceName;
            }
        }
        #endregion

        #region ClassName
        private string _ClassName;
        /// <summary>
        /// ClassName
        /// </summary>
        public string ClassName
        {
            set
            {
                _ClassName = value;
            }
            get
            {
                return _ClassName;
            }
        }
        #endregion

        #region Parameters
        private string _Parameters;
        /// <summary>
        /// Parameters
        /// </summary>
        public string Parameters
        {
            set
            {
                _Parameters = value;
            }
            get
            {
                return _Parameters;
            }
        }
        #endregion

        #region Fields
        private List<Proj_Detail_Field> _Fields = new List<Proj_Detail_Field>();
        /// <summary>
        /// Fields
        /// </summary>
        public List<Proj_Detail_Field> Fields
        {
            set
            {
                _Fields = value;
            }
            get
            {
                return _Fields;
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

        #region 是否保存抓取到的源码文件
        private bool _SaveSourceFile = false;
        /// <summary>
        /// 是否保存抓取到的源码文件
        /// </summary>
        public bool SaveSourceFile
        {
            get
            {
                return _SaveSourceFile;
            }
            set
            {
                _SaveSourceFile = value;
            }
        }
        #endregion

        #region 文件保存位置
        private string _SaveFileDirectory = "";
        /// <summary>
        /// 文件保存位置
        /// </summary>
        public string SaveFileDirectory
        {
            get
            {
                return _SaveFileDirectory;
            }
            set
            {
                _SaveFileDirectory = value;
            }
        }
        #endregion
    }
}
