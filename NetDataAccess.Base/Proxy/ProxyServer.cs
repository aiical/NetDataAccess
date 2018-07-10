using NetDataAccess.Base.Common;
using System;
using System.Collections.Generic;
using System.Net;
using System.Text;

namespace NetDataAccess.Base.Proxy
{
    /// <summary>
    /// 代理服务器信息
    /// </summary>
    public class ProxyServer : IProxyServer
    {
        #region index
        private int _Index = 0;
        public int Index
        {
            get
            {
                return _Index;
            }
            set
            {
                _Index = value;
            }
        }
        #endregion

        #region IP地址
        /// <summary>
        /// IP地址
        /// </summary>
        public string IP { set; get; }
        #endregion

        #region 端口
        /// <summary>
        /// 端口
        /// </summary>
        public int Port { set; get; }
        #endregion

        #region Address
        /// <summary>
        /// Address
        /// </summary>
        public string Address
        {
            get
            {
                return this.IP + ":" + Port.ToString();
            }
        }
        #endregion

        #region 可用时间
        private DateTime _AvailableTime = new DateTime();
        /// <summary>
        /// 可用时间
        /// </summary>
        public DateTime AvailableTime
        {
            get
            {
                return _AvailableTime;
            }
            set
            {
                _AvailableTime = value;
            }
        }
        #endregion

        #region 用户名
        /// <summary>
        /// 用户名
        /// </summary>
        public string User { set; get; }
        #endregion

        #region 密码
        /// <summary>
        /// 密码
        /// </summary>
        public string Pwd { set; get; }
        #endregion

        #region 是否需要用户名密码登录
        /// <summary>
        /// 是否需要用户名密码登录
        /// </summary>
        public bool NeedUserPwd
        {
            get
            {
                return !CommonUtil.IsNullOrBlank(this.User);
            }
        }
        #endregion

        #region 出错次数
        private int _ErrorCount = 0;
        public int ErrorCount
        {
            get
            {
                return _ErrorCount;
            }
            set
            {
                _ErrorCount = value;
            }
        }
        public void AddErrorCount()
        {
            _ErrorCount = _ErrorCount + 1;
        }
        #endregion 

        #region 是否已被弃用
        private bool _IsAbandon = false;
        public bool IsAbandon
        {
            get
            {
                return _IsAbandon;
            }
            set
            {
                _IsAbandon = value;
            }
        }
        #endregion

        #region 构造WebProxy
        /// <summary>
        /// 构造WebProxy
        /// </summary> 
        /// <returns></returns>
        public NdaWebProxy GenerateWebProxy()
        {
            NdaWebProxy wp = new NdaWebProxy(this.Index);
            wp.Address = new Uri("http://" + this.IP + ":" + this.Port.ToString());
            if (!CommonUtil.IsNullOrBlank(this.User))
            {
                wp.Credentials = new NetworkCredential(this.User, this.Pwd);
            }
            return wp;
        }
        #endregion
    }
}
