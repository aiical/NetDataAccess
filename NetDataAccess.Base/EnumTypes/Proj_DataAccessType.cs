using System;
using System.Collections.Generic;
using System.Text;

namespace NetDataAccess.Base.EnumTypes
{
    /// <summary>
    /// 服务器数据获取方式
    /// </summary>
    public enum Proj_DataAccessType
    {
        /// <summary>
        /// 浏览器访问
        /// </summary>
        WebBrowserHtml,

        /// <summary>
        /// 发送request请求Html
        /// </summary>
        WebRequestHtml,

        /// <summary>
        /// 发送request请求Json
        /// </summary>
        WebRequestJson,

        /// <summary>
        /// 发送request请求File
        /// </summary>
        WebRequestFile,

        /// <summary>
        /// 其它方式
        /// </summary>
        OtherAccessType
    }
}
