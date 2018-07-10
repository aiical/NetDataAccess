using System;
using System.Collections.Generic;
using System.Text;

namespace NetDataAccess.Base.EnumTypes
{
    /// <summary>
    /// 加载后动作
    /// </summary>
    public enum ProcessTypeAfterLoaded
    {
        /// <summary>
        /// 无
        /// </summary>
        None,

        /// <summary>
        /// 滚动到底部
        /// </summary>
        ScrollToBottom,

        /// <summary>
        /// 慢慢滚动到底部
        /// </summary>
        ScrollToBottomSlow
    }
}
