using NetDataAccess.Base.EnumTypes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetDataAccess.Base.Definition
{
    public class Proj_CompleteCheck
    {
        #region 文档加载完成检验方式
        private DocumentCompleteCheckType _CheckType = DocumentCompleteCheckType.BrowserCompleteEvent;
        /// <summary>
        /// 文档加载完成检验方式
        /// </summary>
        public DocumentCompleteCheckType CheckType
        {
            get
            {
                return _CheckType;
            }
            set
            {
                _CheckType = value;
            }
        }
        #endregion

        #region 判断值
        private String _CheckValue = "";
        /// <summary>
        /// 判断值，可以使element的xpath、全文匹配字符串、文件大小等
        /// </summary>
        public String CheckValue
        {
            get
            {
                return _CheckValue;
            }
            set
            {
                _CheckValue = value;
            }
        }
        #endregion 
    }
}
