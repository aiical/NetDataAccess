using System;
using System.Collections.Generic;
using System.Text;

namespace NetDataAccess.Base.EnumTypes
{
    /// <summary>
    /// 文档加载完成检验方式
    /// </summary>
    public enum DocumentCompleteCheckType
    {
        /// <summary>
        /// 根据浏览器文档加载完成事件判断
        /// </summary>
        BrowserCompleteEvent,

        /// <summary>
        /// 根据元素是否存在判断（尚未实现）
        /// </summary>
        XpathElementExist,

        /// <summary>
        /// 根据字符串是否存在
        /// </summary>
        TextExist, 

        /// <summary>
        /// 是否以某个字符串结束（检验前需要先Trim）
        /// </summary>
        TrimEndWithText,

        /// <summary>
        /// 以返回的文件大小检验（尚未实现）
        /// </summary>
        FileSize
    }
}
