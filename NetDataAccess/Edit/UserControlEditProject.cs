using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using NetDataAccess.Base.Definition;
using NetDataAccess.Base.Common;
using NetDataAccess.Base.EnumTypes;
using NetDataAccess.Delegate;

namespace NetDataAccess.Edit
{
    /// <summary>
    /// 编辑项目
    /// </summary>
    public partial class UserControlEditProject : UserControl
    { 
        #region 构造函数
        public UserControlEditProject(string groupId, Proj_Main project)
        {
            InitializeComponent();
            this.GroupId = groupId;
            this.Project = project;
            this.comboBoxDetailGrabType.SelectedIndex = 0; 
            this.comboBoxLoginType.SelectedIndex = 1; 
            this.Load += new EventHandler(UserControlEditProject_Load);
        }
        #endregion

        #region Load时显示项目信息
        void UserControlEditProject_Load(object sender, EventArgs e)
        {
            if (this.Project != null)
            {
                this.textBoxName.Text = this.Project.Name;
                this.textBoxDescription.Text = this.Project.Description;
                this.textBoxDetailGrabInfo.Text = this.Project.DetailGrabInfo; 
                this.textBoxLoginPageInfo.Text = this.Project.LoginPageInfo; 
                this.textBoxProgramAfterGrabAll.Text = this.Project.ProgramAfterGrabAll;
                this.textBoxProgramExternalRun.Text = this.Project.ProgramExternalRun;
                this.comboBoxDetailGrabType.Text = this.Project.DetailGrabType.ToString(); 
                this.comboBoxLoginType.Text = this.Project.LoginType.ToString(); 
            }
        }
        #endregion

        #region GroupId
        public string GroupId
        {
            get;
            set;
        }
        #endregion

        #region 项目
        public Proj_Main _Project = null;
        /// <summary>
        /// 项目
        /// </summary>
        public Proj_Main Project 
        {
            get
            {
                return _Project;
            }
            set
            {
                _Project = value;
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
            if (CommonUtil.IsNullOrBlank(this.Project.Id))
            {
                string id = ProjectTaskAccess.AddNewProject(this.Project);
                if (id != null)
                {
                    this.Project.Id = id;
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
               return ProjectTaskAccess.UpdateProject(this.Project);
            }
        }
        #endregion

        #region 验证
        private bool Check()
        {
            string projectName = this.textBoxName.Text.Trim();
            if (CommonUtil.IsNullOrBlank(projectName))
            {
                CommonUtil.Alert("验证", "请录入项目名称.");
                return false;
            }
            else
            {
                //新建项目
                if (this.Project == null)
                {
                    this.Project = new Proj_Main(); 
                    SaveInputToPrjectObject();
                    return true;
                }
                else
                {
                    //判断此项目名称是否存在且Id不同
                    Proj_Main p = ProjectTaskAccess.GetProjectInfoByNameFromDB(projectName);
                    if (p != null && p.Id != this.Project.Id)
                    {
                        CommonUtil.Alert("验证", "已存在的项目名称，请使用其它名称.");
                        return false;
                    }
                    else
                    {
                        SaveInputToPrjectObject();
                        return true;
                    }
                }
            }
        }
        #endregion

        #region 保存录入数据到实体
        private void SaveInputToPrjectObject()
        {
            this.Project.Group_Id = this.GroupId;
            this.Project.Name = this.textBoxName.Text;
            this.Project.Description = this.textBoxDescription.Text;
            this.Project.DetailGrabInfo = this.textBoxDetailGrabInfo.Text; 
            this.Project.LoginPageInfo = this.textBoxLoginPageInfo.Text; 
            this.Project.ProgramAfterGrabAll = this.textBoxProgramAfterGrabAll.Text;
            this.Project.ProgramExternalRun = this.textBoxProgramExternalRun.Text;
            this.Project.DetailGrabType = (DetailGrabType)Enum.Parse(typeof(DetailGrabType), this.comboBoxDetailGrabType.Text); 
            this.Project.LoginType = (LoginLevelType)Enum.Parse(typeof(LoginLevelType), this.comboBoxLoginType.Text); 
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
