using NetDataAccess.Base.Config;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace NetDataAccess.Config
{
    /// <summary>
    /// 系统设置
    /// </summary>
    public partial class FormConfig : Form
    {
        #region 构造函数
        /// <summary>
        /// 构造函数
        /// </summary>
        public FormConfig()
        {
            InitializeComponent();
            this.buttonOK.Click += buttonOK_Click;
            Init();
        }
        #endregion

        #region 点击ok
        void buttonOK_Click(object sender, EventArgs e)
        {
            if (Save())
            {
                this.Close();
            }
        }
        #endregion

        #region 初始化
        private void Init()
        {
            this.checkBoxAllowShowError.Checked = SysConfig.AllowShowError;
        }
        #endregion

        #region 保存
        private bool Save()
        {
            SysConfig.AllowShowError = this.checkBoxAllowShowError.Checked;
            return true;
        }
        #endregion
    }
}
