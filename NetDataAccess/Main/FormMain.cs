using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using NetDataAccess.Base.Definition;
using NetDataAccess.Base.Common;
using NetDataAccess.Base.Proxy;
using System.Diagnostics;
using NetDataAccess.Delegate;
using NetDataAccess.Config;
using NetDataAccess.Edit;
using NetDataAccess.Run; 
using NetDataAccess.Base.DLL;
using NetDataAccess.Base.Config;
using NetDataAccess.Base.Server;
using System.IO;

namespace NetDataAccess.Main
{
    /// <summary>
    /// 主窗口
    /// </summary> 
    public partial class FormMain : Form, IMainRunningContainer
    {
        #region 构造函数
        public FormMain()
        {
            InitializeComponent();

            this.FormClosing += new FormClosingEventHandler(FormMain_FormClosing);
            this.FormClosed += FormMain_FormClosed;
            this.Load += new EventHandler(FormMain_Load);
            this.tabControlMain.MouseDoubleClick += new MouseEventHandler(tabControlMain_MouseDoubleClick);
        }
        #endregion

        #region 双击tabControl
        void tabControlMain_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (this.tabControlMain.SelectedTab is NDATabPage)
            {
                NDATabPage tabPage = (NDATabPage)this.tabControlMain.SelectedTab;
                CloseTabPage(tabPage);
            }
        }
        private void CloseTabPage(NDATabPage tabPage)
        {
            if (tabPage.BeforeClose())
            {
                this.tabControlMain.TabPages.Remove(tabPage);
            }
        }
        #endregion

        #region 刷新窗口名称 added by lixin 20170720
        private void RefreshWindowTitle()
        {
           DirectoryInfo dInfo = Directory.GetParent(Application.StartupPath);
           string title = "NetDataAccess " + dInfo.Name;
           this.Text = title;
        }
        #endregion

        #region 窗口关闭前验证
        void FormMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (MessageBox.Show("确定要关闭吗?", "提示", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                e.Cancel = false;

            }
            else
            {
                e.Cancel = true;
            }
        }

        void FormMain_FormClosed(object sender, FormClosedEventArgs e)
        {
            //关闭Http服务器
            NdaHttpServer.StopServer();

            Process.GetCurrentProcess().Kill();
            System.Environment.Exit(0);
        }
        #endregion

        #region 启动Http服务器
        private void StartHttpServer()
        {
            string ip = SysConfig.ServerIP;
            int port = SysConfig.ServerPort;
            if (ip != null && ip.Length > 0)
            {
                NdaHttpServer.StartServer(ip, port);
            }
        }
        #endregion

        #region LoadSysConfig
        private void LoadSysConfig()
        {
            SysConfig.LoadSysConfig();
        }
        #endregion 

        #region Form Load
        void FormMain_Load(object sender, EventArgs e)
        {
            //加载系统配置
            LoadSysConfig();

            //加载分组数据
            ShowAllGroups();

            //加载项目数据
            ShowAllProjects();  

            //设置运行容器
            SetRunningContainer();

            //启动Http服务器
            StartHttpServer();

            //刷新窗口标题
            RefreshWindowTitle();
        }
        #endregion

        #region 设置运行容器
        private void SetRunningContainer()
        {
            TaskManager.RunningContainer = this;
        }
        #endregion

        #region 显示所有分组、项目
        private void ShowAllGroups()
        {
            List<Proj_Group> allGroups = ProjectTaskAccess.GetAllGroupsFromDB();
            if (allGroups != null)
            {
                foreach (Proj_Group group in allGroups)
                {
                    ShowGroup(group, false);
                }
            }
        }
        private void ShowAllProjects()
        {
            List<Proj_Main> allProjects = ProjectTaskAccess.GetAllProjectsFromDB();
            if (allProjects != null)
            {
                foreach (Proj_Main project in allProjects)
                {
                    TreeNode groupNode = GetGroupNode(project.Group_Id);
                    if (groupNode != null)
                    {
                        ShowProject(groupNode, project, false);
                        RefreshGroupNodeText(groupNode);
                    }
                }
            }
        }

        private TreeNode GetGroupNode(string groupId)
        {
            foreach (TreeNode node in this.treeViewProjectList.Nodes)
            {
                Proj_Group group = (Proj_Group)node.Tag;
                if (group.Id == groupId)
                {
                    return node;
                }
            }
            return null;
        }

        private TreeNode ShowGroup(Proj_Group group, bool checkExist)
        {
            //如果没有
            if (checkExist)
            {
                foreach (TreeNode node in this.treeViewProjectList.Nodes)
                {
                    Proj_Group nodeGroup = (Proj_Group)node.Tag;
                    if (nodeGroup.Id == group.Id)
                    {
                        node.Tag = group;
                        node.Text = group.Name;
                        return node;
                    }
                }
            }

            //如果不是已存在的分组，那么新增一个
            TreeNode newNode = new TreeNode();
            newNode.Tag = group;
            newNode.Text = group.Name;
            this.treeViewProjectList.Nodes.Add(newNode);
            return newNode;
        }
        #endregion

        #region 显示项目
        private TreeNode ShowProject(TreeNode groupNode, Proj_Main project, bool checkExist)
        {
            //如果没有
            if (checkExist)
            {
                foreach (TreeNode node in groupNode.Nodes)
                {
                    Proj_Main nodeProject = (Proj_Main)node.Tag;
                    if (nodeProject.Id == project.Id)
                    {
                        node.Tag = project;
                        node.Text = project.Name;
                        return node;
                    }
                }
            }

            //如果不是已存在的项目，那么新增一个
            TreeNode newNode = new TreeNode();
            newNode.Tag = project;
            newNode.Text = project.Name;
            groupNode.Nodes.Add(newNode);
            return newNode;
        }
        #endregion

        #region 删除分组
        private bool DeleteGroup(Proj_Group group)
        {
            if (CommonUtil.Confirm("确认", "确认删除分组 '"+ group.Name +"' 吗?"))
            {
                return ProjectTaskAccess.DeleteGroup(group.Id);
            }
            else
            {
                return false;
            }
        }
        #endregion

        #region 删除项目
        private bool DeleteProject(Proj_Main project)
        {
            if (CommonUtil.Confirm("确认", "确认删除项目 '" + project.Name + "' 吗?"))
            {
                return ProjectTaskAccess.DeleteProject(project.Id);
            }
            else
            {
                return false;
            }
        }
        #endregion

        #region 查找TabPage
        private NDATabPage CheckExistTabPage(string id)
        {
            foreach (TabPage tp in this.tabControlMain.TabPages)
            {
                if (tp is NDATabPage)
                {
                    NDATabPage tabPage = (NDATabPage)tp;
                    if (tabPage.Id == id)
                    {
                        return tabPage;
                    }
                }
            }
            return null;
        }
        #endregion

        #region 新建TabPage
        private NDATabPage CreateTabPage(string id, string text)
        {
            NDATabPage newTabPage = new NDATabPage();
            newTabPage.Id = id;
            newTabPage.Text = text;
            this.tabControlMain.TabPages.Add(newTabPage);
            return newTabPage;
        }
        #endregion

        #region 新建分组按钮
        private void toolStripMenuItemAddGroup_Click(object sender, EventArgs e)
        {
            string id = "New_Group";
            NDATabPage tabPage = CheckExistTabPage(id);
            if (tabPage == null)
            {
                UserControlEditGroup groupControl = new UserControlEditGroup(null);
                groupControl.AfterSaveEvent += new AfterSaveHandler(groupControl_AfterSaveEvent);
                groupControl.Dock = DockStyle.Fill;
                tabPage = this.CreateTabPage(id, "新建分组");
                tabPage.Controls.Add(groupControl);
                tabPage.BeforeTabPageCloseEvent += new BeforeTabPageCloseHandler(tabPage_GroupBeforeTabPageCloseEvent);
            }
            this.tabControlMain.SelectedTab = tabPage;

        }

        bool tabPage_GroupBeforeTabPageCloseEvent(object sender, EventArgs e)
        {
            NDATabPage tabPage = (NDATabPage)sender;
            UserControlEditGroup groupControl = (UserControlEditGroup)tabPage.Controls[0];
            return groupControl.Close();
        } 

        void groupControl_AfterSaveEvent(object sender, EventArgs e)
        {
            UserControlEditGroup groupControl = (UserControlEditGroup)sender;
            TreeNode groupNode = ShowGroup(groupControl.Group, true);
            RefreshGroupNodeText(groupNode);
            this.treeViewProjectList.SelectedNode = groupNode;
            NDATabPage tabPage = (NDATabPage)groupControl.Parent;
            tabPage.Id = groupControl.Group.Id;
            tabPage.Text = "分组:" + groupControl.Group.Name;
        }
        #endregion

        #region 新建项目按钮
        private void toolStripMenuItemAddProject_Click(object sender, EventArgs e)
        {
            TreeNode selectedNode = this.treeViewProjectList.SelectedNode;
            if (selectedNode == null)
            {
                CommonUtil.Alert("提示", "请选择或新建分组.");
            }
            else
            {
                TreeNode groupNode = ((selectedNode.Tag is Proj_Main) ? selectedNode.Parent : selectedNode);
                Proj_Group group = (Proj_Group)groupNode.Tag;

                string id = group.Id + "_New_Project";
                NDATabPage tabPage = CheckExistTabPage(id);
                if (tabPage == null)
                {
                    UserControlEditProject projectControl = new UserControlEditProject(group.Id, null);
                    projectControl.AfterSaveEvent += new AfterSaveHandler(projectControl_AfterSaveEvent);
                    projectControl.Dock = DockStyle.Fill;
                    tabPage = this.CreateTabPage(id, "新建项目");
                    tabPage.Controls.Add(projectControl);
                    tabPage.BeforeTabPageCloseEvent += new BeforeTabPageCloseHandler(tabPage_ProjectBeforeTabPageCloseEvent);
                }
                this.tabControlMain.SelectedTab = tabPage;
            }
        }

        bool tabPage_ProjectBeforeTabPageCloseEvent(object sender, EventArgs e)
        {
            NDATabPage tabPage = (NDATabPage)sender;
            UserControlEditProject projectControl = (UserControlEditProject)tabPage.Controls[0];
            return projectControl.Close();
        }

        void projectControl_AfterSaveEvent(object sender, EventArgs e)
        {
            UserControlEditProject projectControl = (UserControlEditProject)sender;
            TreeNode groupNode = GetGroupNode(projectControl.GroupId);
            if (groupNode != null)
            {
                TreeNode projectNode = ShowProject(groupNode, projectControl.Project, true);
                RefreshGroupNodeText(groupNode);
                groupNode.Expand();
                this.treeViewProjectList.SelectedNode = projectNode;
                NDATabPage tabPage = (NDATabPage)projectControl.Parent;
                tabPage.Id = projectControl.Project.Id;
                tabPage.Text = "项目:" + projectControl.Project.Name;
            }
            else
            {
                CommonUtil.Alert("提示", "分组已被删除.");
            }
        }
        #endregion 

        #region 删除按钮
        private void toolStripButtonDelete_Click(object sender, EventArgs e)
        {
            TreeNode selectedNode = this.treeViewProjectList.SelectedNode;
            if (selectedNode == null)
            {
                CommonUtil.Alert("提示", "无选中项");
            }
            else
            {
                if (selectedNode.Tag is Proj_Group)
                {
                    if (DeleteGroup((Proj_Group)selectedNode.Tag))
                    {
                        selectedNode.Remove();
                    }
                }
                else if (selectedNode.Tag is Proj_Main)
                { 
                    if (DeleteProject((Proj_Main)selectedNode.Tag))
                    {
                        TreeNode groupNode = selectedNode.Parent;
                        selectedNode.Remove();
                        RefreshGroupNodeText(groupNode);
                    }
                }
            }
        }
        #endregion

        #region 获取节点
        private Proj_Main GetProject(string groupName, string projectName)
        {
            TreeNodeCollection allGroupNodes = this.treeViewProjectList.Nodes;
            if (allGroupNodes != null)
            {
                foreach (TreeNode groupNode in allGroupNodes)
                {
                    Proj_Group group = groupNode.Tag as Proj_Group;
                    if (group.Name == groupName)
                    {
                        TreeNodeCollection allProjectNodes = groupNode.Nodes;
                        if (allProjectNodes != null)
                        {
                            foreach (TreeNode projectNode in allProjectNodes)
                            {
                                Proj_Main proj = projectNode.Tag as Proj_Main;
                                if (proj.Name == projectName)
                                {
                                    return proj;
                                }
                            }
                        }
                    }
                }
            }
            return null;
        }
        #endregion

        #region 自动执行
        public void InvokeRunTask(string groupName, string projectName, string listFilePath, string inputDir, string middleDir, string outputDir, string parameters, string stepId, bool autoRun, bool popPrompt)
        {
            this.Invoke(new RunDelegate(RunTask), new object[] { groupName, projectName, listFilePath, inputDir, middleDir, outputDir,  parameters, stepId, autoRun, popPrompt });
        }
        private delegate void RunDelegate(string groupName, string projectName, string listFilePath, string inputDir, string middleDir, string outputDir, string parameters, string stepId, bool autoRun, bool popPrompt);
        private void RunTask(string groupName, string projectName, string listFilePath, string inputDir, string middleDir, string outputDir, string parameters, string stepId, bool autoRun, bool popPrompt)
        {
            Proj_Main project = GetProject(groupName, projectName);
            if (project != null)
            { 
                string tabId = autoRun ? stepId : projectName;
                NDATabPage tabPage = CheckExistTabPage(tabId);
                UserControlRunGrabWebPage runControl = new UserControlRunGrabWebPage(project, autoRun, popPrompt, listFilePath, inputDir, middleDir, outputDir, parameters, stepId);
                runControl.Dock = DockStyle.Fill;
                tabPage = this.CreateTabPage(tabId, "执行Task step:" + projectName + "(stepId=" + stepId + ")");
                tabPage.Controls.Add(runControl);
                tabPage.BeforeTabPageCloseEvent += new BeforeTabPageCloseHandler(tabPage_RunBeforeTabPageCloseEvent);
                this.tabControlMain.SelectedTab = tabPage;
            }
            else
            {
                throw new Exception("不存在的项目, groupName = " + groupName + ", projectName = " + projectName + ".");
            }
        }
        #endregion

        #region 关闭任务UI
        public void InvokeCloseTaskUI(string taskId)
        {
            this.Invoke(new CloseTaskUIDelegate(CloseTaskUI), new object[] { taskId });
        }
        private delegate void CloseTaskUIDelegate(string taskId);
        #endregion

        #region 关闭任务UI
        private void CloseTaskUI(string taskId)
        {
            foreach (TabPage tabPage in this.tabControlMain.TabPages)
            {
                if (tabPage is NDATabPage)
                {
                    NDATabPage ndaTabPage = (NDATabPage)tabPage;
                    if (ndaTabPage.Id == taskId)
                    {
                        CloseTabPage(ndaTabPage);
                        break;
                    }
                }
            }
        }
        #endregion

        #region 执行按钮
        private void toolStripButtonRun_Click(object sender, EventArgs e)
        {
            TreeNode selectedNode = this.treeViewProjectList.SelectedNode;
            if (selectedNode.Tag is Proj_Main)
            {
                Proj_Main project =(Proj_Main)selectedNode.Tag ;
                string tabId = "Run_" + project.Id;
                NDATabPage tabPage = CheckExistTabPage(tabId);
                if (tabPage == null)
                {
                    UserControlRunGrabWebPage runControl = new UserControlRunGrabWebPage(project);
                    runControl.Dock = DockStyle.Fill;
                    tabPage = this.CreateTabPage(tabId, "执行:" + project.Name);
                    tabPage.Controls.Add(runControl);
                    tabPage.BeforeTabPageCloseEvent += new BeforeTabPageCloseHandler(tabPage_RunBeforeTabPageCloseEvent);
                }
                this.tabControlMain.SelectedTab = tabPage;
            }
        } 
        bool tabPage_RunBeforeTabPageCloseEvent(object sender, EventArgs e)
        {
            NDATabPage tabPage = (NDATabPage)sender;
            UserControlRunGrabWebPage runControl = (UserControlRunGrabWebPage)tabPage.Controls[0];
            return runControl.Close();
        }
        #endregion

        #region 编辑按钮
        private void toolStripButtonEdit_Click(object sender, EventArgs e)
        {
            TreeNode selectedNode = this.treeViewProjectList.SelectedNode;
            EditNodeObject(selectedNode);
        }
        #endregion

        #region 编辑节点对应的对象
        private void EditNodeObject(TreeNode selectedNode)
        {
            if (selectedNode == null)
            {
                CommonUtil.Alert("提示", "无选中项");
            }
            else
            {
                if (selectedNode.Tag is Proj_Group)
                {
                    Proj_Group group = (Proj_Group)selectedNode.Tag;

                    string id = group.Id;
                    NDATabPage tabPage = CheckExistTabPage(id);
                    if (tabPage == null)
                    {
                        UserControlEditGroup groupControl = new UserControlEditGroup(group);
                        groupControl.AfterSaveEvent += new AfterSaveHandler(groupControl_AfterSaveEvent);
                        groupControl.Dock = DockStyle.Fill;
                        tabPage = this.CreateTabPage(id, "分组:" + group.Name);
                        tabPage.Controls.Add(groupControl);
                        tabPage.BeforeTabPageCloseEvent += new BeforeTabPageCloseHandler(tabPage_GroupBeforeTabPageCloseEvent);
                    }
                    this.tabControlMain.SelectedTab = tabPage;
                }
                else if (selectedNode.Tag is Proj_Main)
                {
                    Proj_Main project = (Proj_Main)selectedNode.Tag;
                    Proj_Group group = (Proj_Group)selectedNode.Parent.Tag;

                    string id = project.Id;
                    NDATabPage tabPage = CheckExistTabPage(id);
                    if (tabPage == null)
                    {
                        UserControlEditProject projectControl = new UserControlEditProject(group.Id, project);
                        projectControl.AfterSaveEvent += new AfterSaveHandler(projectControl_AfterSaveEvent);
                        projectControl.Dock = DockStyle.Fill;
                        tabPage = this.CreateTabPage(id, "项目:" + project.Name);
                        tabPage.Controls.Add(projectControl);
                        tabPage.BeforeTabPageCloseEvent += new BeforeTabPageCloseHandler(tabPage_ProjectBeforeTabPageCloseEvent);
                    }
                    this.tabControlMain.SelectedTab = tabPage;
                }
            }

        }
        #endregion 
        
        #region 刷新分组节点文字
        private void RefreshGroupNodeText(TreeNode groupNode)
        {
            Proj_Group group = (Proj_Group)groupNode.Tag;
            groupNode.Text = group.Name + "(" + groupNode.Nodes.Count.ToString() + ")";
        }
        #endregion

        #region 配置按钮
        private void configToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormConfig configForm = new FormConfig();
            configForm.ShowDialog(this);
        }
        #endregion

        #region About
        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Net Data Access 2.0", "关于");
        }
        #endregion 

        #region 检测对外服务接口是否正常提供 added by lixin 20170720
        private void timerServer_Tick(object sender, EventArgs e)
        {
            bool succeed = NdaHttpServer.TestServer();
            this.toolStripStatusLabelServerStatus.Text = "服务接口状态(" + SysConfig.ServerIP + ":" + SysConfig.ServerPort.ToString() + "): " + (succeed ? "正常" : "无服务");
        }
        #endregion 
    }
}
