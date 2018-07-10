using NetDataAccess.Base.EnumTypes;
using System;
using System.Collections.Generic;
using System.Text;

namespace NetDataAccess.Base.Definition
{
    /// <summary>
    /// 各种爬取环节定义的interface
    /// </summary>
    public interface IProj_XmlConfig
    {
        bool NeedProxy { get; set; }
        bool AutoAbandonDisableProxy { get; set; }

    }
}
