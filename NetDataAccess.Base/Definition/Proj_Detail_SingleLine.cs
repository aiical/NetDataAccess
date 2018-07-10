using NetDataAccess.Base.EnumTypes;
using System;
using System.Collections.Generic;
using System.Text;

namespace NetDataAccess.Base.Definition
{
    /// <summary>
    /// 列表页设置
    /// </summary>
    public class Proj_Detail_SingleLine : IProj_FieldConfig
    { 
        #region IntervalAfterLoaded
        private decimal _IntervalAfterLoaded;
        /// <summary>
        /// IntervalAfterLoaded
        /// </summary>
        public decimal IntervalAfterLoaded
        {
            set
            {
                _IntervalAfterLoaded = value;
            }
            get
            {
                return _IntervalAfterLoaded;
            }
        }
        #endregion

        #region IntervalProxyRequest
        private int _IntervalProxyRequest = 0;
        /// <summary>
        /// 同一个代理，两次访问的间隔时间
        /// </summary>
        public int IntervalProxyRequest
        {
            set
            {
                _IntervalProxyRequest = value;
            }
            get
            {
                return _IntervalProxyRequest;
            }
        }
        #endregion

        #region html获取方式
        private Proj_DataAccessType _DataAccessType = Proj_DataAccessType.WebBrowserHtml;
        /// <summary>
        /// html获取方式
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

        #region 获取多少页详情页数据后保存一次
        private int _IntervalDetailPageSave = 10;
        /// <summary>
        /// 获取多少页详情页数据后保存一次
        /// </summary>
        public int IntervalDetailPageSave
        {
            get
            {
                return _IntervalDetailPageSave;
            }
            set
            {
                _IntervalDetailPageSave = value;
            }
        }
        #endregion

        #region 获取页码范围，抓取的起始页码
        private int _StartPageIndex = 0;
        /// <summary>
        /// 抓取的起始页码
        /// </summary>
        public int StartPageIndex
        {
            get
            {
                return _StartPageIndex;
            }
            set
            {
                _StartPageIndex = value;
            }
        }
        #endregion

        #region 获取页码范围，抓取的起始页码
        private int _EndPageIndex = 0;
        /// <summary>
        /// 抓取的结束页码
        /// </summary>
        public int EndPageIndex
        {
            get
            {
                return _EndPageIndex;
            }
            set
            {
                _EndPageIndex = value;
            }
        }
        #endregion

        #region 详情页信息输出类型
        private ExportType _ExportType = ExportType.Excel;
        /// <summary>
        /// 详情页信息输出类型
        /// </summary>
        public ExportType ExportType
        {
            get
            {
                return _ExportType;
            }
            set
            {
                _ExportType = value;
            }
        }
        #endregion

        #region 允许自动放弃某条任务
        private bool _AllowAutoGiveUp = false;
        /// <summary>
        /// 允许自动放弃某条任务
        /// </summary>
        public bool AllowAutoGiveUp
        {
            get
            {
                return _AllowAutoGiveUp;
            }
            set
            {
                _AllowAutoGiveUp = value;
            }
        }
        #endregion

        #region 需要拆分不同文件夹存储
        private bool _NeedPartDir = false;
        /// <summary>
        /// 需要拆分不同文件夹存储
        /// </summary>
        public bool NeedPartDir
        {
            get
            {
                return _NeedPartDir;
            }
            set
            {
                _NeedPartDir = value;
            }
        }
        #endregion

        #region 同时抓取的线程数
        private int _ThreadCount = 5;
        /// <summary>
        /// 同时抓取的线程数
        /// </summary>
        public int ThreadCount
        {
            get
            {
                return _ThreadCount;
            }
            set
            {
                _ThreadCount = value;
            }
        }
        #endregion

        #region 请求超时时间（ms）
        private int _RequestTimeout = 30 * 1000;
        /// <summary>
        /// 请求超时时间（ms）
        /// </summary>
        public int RequestTimeout
        {
            get
            {
                return _RequestTimeout;
            }
            set
            {
                _RequestTimeout = value;
            }
        }
        #endregion

        #region 编码方式
        private String _Encoding = null;
        /// <summary>
        /// 编码方式
        /// </summary>
        public String Encoding
        {
            get
            {
                return _Encoding;
            }
            set
            {
                _Encoding = value;
            }
        }
        #endregion

        #region 请求串头文件x-requested-with
        private String _XRequestedWith = "";
        /// <summary>
        /// 请求串头文件x-requested-with
        /// </summary>
        public String XRequestedWith
        {
            get
            {
                return _XRequestedWith;
            }
            set
            {
                _XRequestedWith = value;
            }
        }
        #endregion

        #region 文档加载完成检验方式
        private Proj_CompleteCheckList _CompleteChecks;
        /// <summary>
        /// 文档加载完成检验方式
        /// </summary>
        public Proj_CompleteCheckList CompleteChecks
        {
            get
            {
                return _CompleteChecks;
            }
            set
            {
                _CompleteChecks = value;
            }
        }
        #endregion 
    }
}
