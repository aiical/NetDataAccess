using NetDataAccess.Base.EnumTypes;
using System;
using System.Collections.Generic;
using System.Text;

namespace NetDataAccess.Base.Definition
{
    /// <summary>
    /// 字段
    /// </summary>
    public interface IProj_FieldConfig : IProj_XmlConfig
    {
        #region Fields
        /// <summary>
        /// Fields
        /// </summary>
        List<Proj_Detail_Field> Fields { get; }
        #endregion
    }
}
