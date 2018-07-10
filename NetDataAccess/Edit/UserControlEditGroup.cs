using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using NetDataAccess.Base.Definition;
using NetDataAccess.Base.Common;
using NetDataAccess.Delegate;

namespace NetDataAccess.Edit
{
    /// <summary>
    /// 编辑Group
    /// </summary>
    public partial class UserControlEditGroup : UserControl
    { 
        #region 构造函数
        public UserControlEditGroup(Proj_Group group)
        {
            InitializeComponent();
            this.Group = group;
            this.Load += new EventHandler(UserControlEditGroup_Load);
        }
        #endregion

        #region Load时显示分组信息
        private void UserControlEditGroup_Load(object sender, EventArgs e)
        {
            if (this.Group != null)
            {
                this.textBoxName.Text = this.Group.Name;
                this.textBoxDescription.Text = this.Group.Description;
            }
        }
        #endregion

        #region 分组
        public Proj_Group _Group = null;
        /// <summary>
        /// 分组
        /// </summary>
        public Proj_Group Group 
        {
            get
            {
                return _Group;
            }
            set
            {
                _Group = value;
            }
        }
        #endregion

        #region 保存后事件
        public event AfterSaveHandler AfterSaveEvent;
        #endregion 

        #region 点击确认按钮
        private void buttonOK_Click(object sender, EventArgs e)
        {
            Save();
        }
        #endregion

        #region 保存
        public bool Save()
        {
            if (Check() && SaveToDB())
            {
                if (this.AfterSaveEvent != null)
                {
                    this.AfterSaveEvent(this, new EventArgs());
                }
                return true;
            }
            else
            {
                return false;
            }
        }
        #endregion

        #region 保存
        private bool SaveToDB()
        {
            if (CommonUtil.IsNullOrBlank(this.Group.Id))
            {
                string id = ProjectTaskAccess.AddNewGroup(this.Group);
                if (id != null)
                {
                    this.Group.Id = id;
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
               return ProjectTaskAccess.UpdateGroup(this.Group);
            }
        }
        #endregion

        #region 验证
        private bool Check()
        {
            string groupName = this.textBoxName.Text.Trim();
            if (CommonUtil.IsNullOrBlank(groupName))
            {
                CommonUtil.Alert("验证", "请录入组名.");
                return false;
            }
            else
            {
                //新建分组
                if (this.Group == null)
                {
                    this.Group = new Proj_Group();
                    this.Group.Name = groupName;
                    this.Group.Description = this.textBoxDescription.Text.Trim();
                    return true;
                }
                else
                {
                    //判断此组名是否存在且Id不同
                    Proj_Group g = ProjectTaskAccess.GetGroupInfoByNameFromDB(groupName);
                    if (g != null && g.Id != this.Group.Id)
                    {
                        CommonUtil.Alert("验证", "已存在的组名，请使用其它名称.");
                        return false;
                    }
                    else
                    {
                        this.Group.Name = groupName;
                        this.Group.Description = this.textBoxDescription.Text.Trim();
                        return true;
                    }
                }
            }
        }
        #endregion

        #region 关闭
        public bool Close()
        {
            switch (MessageBox.Show("关闭前执行保存吗?", "确认", MessageBoxButtons.YesNoCancel))
            {
                case DialogResult.Yes:
                    return Save();
                case DialogResult.No:
                    return true;
                case DialogResult.Cancel:
                default:
                    return false;
            }
        }
        
        private void buttonClose_Click(object sender, EventArgs e)
        {
            if (this.Close())
            {
                TabPage parentTabPage = (TabPage)this.Parent;
                ((TabControl)parentTabPage.Parent).TabPages.Remove(parentTabPage);
            }
        }
        #endregion
    }
}
