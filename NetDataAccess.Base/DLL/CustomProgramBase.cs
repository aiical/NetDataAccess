
using NetDataAccess.Base.UI;
using System;
using System.Collections.Generic;
using System.Text;

namespace NetDataAccess.Base.DLL
{
    /// <summary>
    /// 自定义程序基类
    /// </summary>
    [System.Runtime.InteropServices.ComVisibleAttribute(true)]
    public class CustomProgramBase
    {
        #region RunPage
        private IRunWebPage _RunPage = null;
        protected IRunWebPage RunPage
        {
            get
            {
                return _RunPage;
            }
        }
        #endregion

        #region 初始化
        public void Init(IRunWebPage runPage)
        {
            this._RunPage = runPage;
        }
        #endregion
        
        #region 运行
        public virtual bool Run(string parameters)
        {
            return false;
        }
        #endregion
    }
}
