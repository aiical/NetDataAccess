using NetDataAccess.Delegate;
using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;

namespace NetDataAccess.Main
{
    /// <summary>
    /// Tab页
    /// </summary>
    public class NDATabPage:TabPage
    {
        #region Id
        public string Id
        {
            get; 
            set; 
        }
        #endregion

        #region 关闭前事件
        public event BeforeTabPageCloseHandler BeforeTabPageCloseEvent;
        #endregion

        #region 关闭前验证
        /// <summary>
        /// 关闭前验证
        /// </summary>
        /// <returns></returns>
        public bool BeforeClose()
        {
            if (this.BeforeTabPageCloseEvent != null)
            {
                return this.BeforeTabPageCloseEvent(this, new EventArgs());
            }
            else
            {
                return false;
            }
        }
        #endregion
    }
}
