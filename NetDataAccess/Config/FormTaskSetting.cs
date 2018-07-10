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
    /// 任务设置
    /// </summary>
    public partial class FormTaskSetting : Form
    {
        #region 构造函数
        /// <summary>
        /// 构造函数
        /// </summary>
        public FormTaskSetting()
        {
            InitializeComponent();
            this.buttonOK.Click += buttonOK_Click;
            Init();
        }
        #endregion

        #region 记录设置结果
        private Dictionary<string, object> _Setting = null;
        public Dictionary<string, object> Setting
        {
            get
            {
                return _Setting;
            }
        }
        #endregion

        #region 按钮ok
        void buttonOK_Click(object sender, EventArgs e)
        {
            if (Save())
            {
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
        }
        #endregion

        #region 初始化
        private void Init()
        { 
            this.checkBoxRunGrabDetail.Checked = true;
            this.checkBoxRunRead.Checked = true;
            this.checkBoxRunExport.Checked = true;
            this.checkBoxRunCustom.Checked = true;
        }
        #endregion

        #region 保存设置结果
        private bool Save()
        {
            _Setting = new Dictionary<string, object>(); 
            _Setting.Add("RunGrabDetail", this.checkBoxRunGrabDetail.Checked);
            _Setting.Add("RunRead", this.checkBoxRunRead.Checked);
            _Setting.Add("RunExport", this.checkBoxRunExport.Checked);
            _Setting.Add("RunCustom", this.checkBoxRunCustom.Checked);
            return true;
        }
        #endregion

        #region 取消
        private void buttonCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
        #endregion
    }
}
