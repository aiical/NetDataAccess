using System;
using System.Collections.Generic;
using System.Text;

namespace NetDataAccess.Base.EnumTypes
{
    /// <summary>
    /// 详情页数据的处理方法
    /// </summary>
    public enum DetailGrabType
    {
        /// <summary>
        /// 按照单条数据处理，使用绝对路径定位各个元素，然后记录各个元素的值
        /// </summary>
        SingleLineType,

        /// <summary>
        /// 需要先定位分行控件找到多行，然后逐行找到各个字段值
        /// </summary>
        MultiLineType,

        /// <summary>
        /// 无详情页
        /// </summary>
        NoneDetailPage,

        /// <summary>
        /// 自定义代码处理
        /// </summary>
        ProgramType
    }
}
