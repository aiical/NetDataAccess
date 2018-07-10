using NetDataAccess.Base.Config;
using NetDataAccess.Base.Proxy;
using System;
using System.Collections.Generic;
using System.Net;
using System.Text;

namespace NetDataAccess.Base.Web
{
    public class NDAWebClient : WebClient
    {
        private ProxyServer _ProxyServer = null;
        public ProxyServer ProxyServer
        {
            get
            {
                return this._ProxyServer;
            }
            set
            {
                this._ProxyServer = value;
                this.Proxy = ProxyServer.GenerateWebProxy();
            }
        }

        private string _Id = null;
        public string Id
        {
            get
            {
                return _Id;
            }
            set
            {
                _Id = value;
            }
        }

        private Encoding _ResponseEncoding = null;
        public Encoding ResponseEncoding
        {
            get
            {
                return _ResponseEncoding;
            }
            set
            {
                _ResponseEncoding = value;
            }
        } 

        /// <summary>
        /// 获取网络请求
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        protected override WebRequest GetWebRequest(Uri address)
        {
            WebRequest wq = base.GetWebRequest(address);
            wq.Timeout = this.Timeout;
            return wq;
        }

        private int _Timeout = SysConfig.WebPageRequestTimeout;
        /// <summary>
        /// 超时时间（毫秒）
        /// </summary>
        public int Timeout
        {
            get
            {
                return this._Timeout;
            }
            set
            {
                this._Timeout = value;
            }
        }
    }
}
