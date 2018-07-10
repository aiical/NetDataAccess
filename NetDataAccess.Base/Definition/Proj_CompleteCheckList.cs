using NetDataAccess.Base.EnumTypes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetDataAccess.Base.Definition
{
    public class Proj_CompleteCheckList : List<Proj_CompleteCheck>
    {
        #region 条件为并且
        private bool _AndCondition = false;
        /// <summary>
        /// 条件为并且
        /// </summary>
        public bool AndCondition
        {
            get
            {
                return _AndCondition;
            }
            set
            {
                _AndCondition = value;
            }
        }
        #endregion
    }
}
