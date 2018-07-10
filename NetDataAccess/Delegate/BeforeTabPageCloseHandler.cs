using System;
using System.Collections.Generic;
using System.Text;

namespace NetDataAccess.Delegate
{
    /// <summary>
    /// tab页关闭前
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    /// <returns></returns>
    public delegate bool BeforeTabPageCloseHandler(object sender, System.EventArgs e);
}