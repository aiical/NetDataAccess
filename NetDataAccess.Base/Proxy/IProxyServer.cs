using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetDataAccess.Base.Proxy
{
    public interface IProxyServer
    {
        #region 构造WebProxy
        /// <summary>
        /// 构造WebProxy
        /// </summary> 
        /// <returns></returns>
        NdaWebProxy GenerateWebProxy();
        #endregion

        #region index
        int Index { get; }
        #endregion

        #region IP地址
        /// <summary>
        /// IP地址
        /// </summary>
        string IP { get; }
        #endregion

        #region 端口
        /// <summary>
        /// 端口
        /// </summary>
        int Port { get; }
        #endregion

        #region Address
        /// <summary>
        /// Address
        /// </summary>
        string Address { get; }
        #endregion

        #region 用户名
        /// <summary>
        /// 用户名
        /// </summary>
        string User { get; }
        #endregion

        #region 密码
        /// <summary>
        /// 密码
        /// </summary>
        string Pwd { get; }
        #endregion

        #region 是否需要用户名密码登录
        /// <summary>
        /// 是否需要用户名密码登录
        /// </summary>
        bool NeedUserPwd { get; }
        #endregion

        #region 出错次数
        int ErrorCount { get; }
        void AddErrorCount();
        #endregion 

        #region 是否已被弃用
        bool IsAbandon { get; }
        #endregion

        #region 可用时间
        /// <summary>
        /// 可用时间
        /// </summary>
        DateTime AvailableTime { get; }
        #endregion
    }
}
