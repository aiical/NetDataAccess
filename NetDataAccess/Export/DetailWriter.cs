using NetDataAccess.Base.DB;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NetDataAccess.Export
{
    /// <summary>
    /// 文档输出接口
    /// </summary>
    internal interface DetailExportWriter
    {
        #region 保存一条记录
        /// <summary>
        /// 保存一条记录
        /// </summary>
        /// <param name="listSheet"></param>
        /// <param name="columnNameToIndex"></param>
        /// <param name="fieldValues"></param>
        /// <param name="rowIndex"></param>
        /// <param name="pageUrl"></param>
        void SaveDetailFieldValue(IListSheet listSheet, Dictionary<string, int> columnNameToIndex, Dictionary<string, string> fieldValues, int rowIndex, string pageUrl);
        #endregion

        #region 保存到硬盘
        /// <summary>
        /// 保存到硬盘
        /// </summary>
        void SaveToDisk();
        #endregion
    }
}
