using NetDataAccess.Base.EnumTypes;
using System;
using System.Collections.Generic;
using System.Text;

namespace NetDataAccess.Base.Definition
{
    /// <summary>
    /// 字段设置
    /// </summary>
    public class Proj_Detail_Field
    {
        #region Name
        private string _Name;
        /// <summary>
        /// Name
        /// </summary>
        public string Name
        {
            set
            {
                _Name = value;
            }
            get
            {
                return _Name;
            }
        }
        #endregion

        #region Path
        private string _Path;
        /// <summary>
        /// Path
        /// </summary>
        public string Path
        {
            set
            {
                _Path = value;
            }
            get
            {
                return _Path;
            }
        }
        #endregion

        #region IsAbsolute
        private bool _IsAbsolute;
        /// <summary>
        /// IsAbsolute
        /// </summary>
        public bool IsAbsolute
        {
            set
            {
                _IsAbsolute = value;
            }
            get
            {
                return _IsAbsolute;
            }
        }
        #endregion

        #region AttributeName
        private string _AttributeName = "";
        /// <summary>
        /// AttributeName
        /// </summary>
        public string AttributeName
        {
            set
            {
                _AttributeName = value;
            }
            get
            {
                return _AttributeName;
            }
        }
        #endregion

        #region 是否记录下整个outhtml
        private bool _NeedAllHtml = false;
        /// <summary>
        /// 是否记录下整个outhtml
        /// </summary>
        public bool NeedAllHtml
        {
            set
            {
                _NeedAllHtml = value;
            }
            get
            {
                return _NeedAllHtml;
            }
        }
        #endregion

        #region Type
        private FieldValueType _Type = FieldValueType.String;
        /// <summary>
        /// AttributeName
        /// </summary>
        public FieldValueType Type
        {
            set
            {
                _Type = value;
            }
            get
            {
                return _Type;
            }
        }
        #endregion

        #region ColumnWidth 字符个数
        private int _ColumnWidth = 10;
        /// <summary>
        /// ColumnWidth 字符个数
        /// </summary>
        public int ColumnWidth
        {
            set
            {
                _ColumnWidth = value;
            }
            get
            {
                return _ColumnWidth;
            }
        }
        #endregion  
    }
}
