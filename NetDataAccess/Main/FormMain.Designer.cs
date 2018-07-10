namespace NetDataAccess.Main
{
    partial class FormMain
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormMain));
            this.splitContainerMain = new System.Windows.Forms.SplitContainer();
            this.panelLeftCenter = new System.Windows.Forms.Panel();
            this.treeViewProjectList = new System.Windows.Forms.TreeView();
            this.panelLeftTop = new System.Windows.Forms.Panel();
            this.toolStripProjectList = new System.Windows.Forms.ToolStrip();
            this.toolStripDropDownButtonAdd = new System.Windows.Forms.ToolStripDropDownButton();
            this.toolStripMenuItemAddGroup = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItemAddProject = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripButtonEdit = new System.Windows.Forms.ToolStripButton();
            this.toolStripButtonDelete = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.toolStripButtonRun = new System.Windows.Forms.ToolStripButton();
            this.tabControlMain = new System.Windows.Forms.TabControl();
            this.tabPageMain = new System.Windows.Forms.TabPage();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.settingToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.configToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.helpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.aboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.timerServer = new System.Windows.Forms.Timer(this.components);
            this.statusStripMain = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabelServerStatus = new System.Windows.Forms.ToolStripStatusLabel();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainerMain)).BeginInit();
            this.splitContainerMain.Panel1.SuspendLayout();
            this.splitContainerMain.Panel2.SuspendLayout();
            this.splitContainerMain.SuspendLayout();
            this.panelLeftCenter.SuspendLayout();
            this.panelLeftTop.SuspendLayout();
            this.toolStripProjectList.SuspendLayout();
            this.tabControlMain.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.statusStripMain.SuspendLayout();
            this.SuspendLayout();
            // 
            // splitContainerMain
            // 
            this.splitContainerMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainerMain.FixedPanel = System.Windows.Forms.FixedPanel.Panel1;
            this.splitContainerMain.Location = new System.Drawing.Point(0, 25);
            this.splitContainerMain.Name = "splitContainerMain";
            // 
            // splitContainerMain.Panel1
            // 
            this.splitContainerMain.Panel1.Controls.Add(this.panelLeftCenter);
            this.splitContainerMain.Panel1.Controls.Add(this.panelLeftTop);
            this.splitContainerMain.Panel1.Padding = new System.Windows.Forms.Padding(1);
            // 
            // splitContainerMain.Panel2
            // 
            this.splitContainerMain.Panel2.Controls.Add(this.tabControlMain);
            this.splitContainerMain.Size = new System.Drawing.Size(1012, 507);
            this.splitContainerMain.SplitterDistance = 241;
            this.splitContainerMain.TabIndex = 1;
            // 
            // panelLeftCenter
            // 
            this.panelLeftCenter.Controls.Add(this.treeViewProjectList);
            this.panelLeftCenter.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelLeftCenter.Location = new System.Drawing.Point(1, 26);
            this.panelLeftCenter.Name = "panelLeftCenter";
            this.panelLeftCenter.Size = new System.Drawing.Size(239, 480);
            this.panelLeftCenter.TabIndex = 0;
            // 
            // treeViewProjectList
            // 
            this.treeViewProjectList.Dock = System.Windows.Forms.DockStyle.Fill;
            this.treeViewProjectList.Font = new System.Drawing.Font("宋体", 10F);
            this.treeViewProjectList.FullRowSelect = true;
            this.treeViewProjectList.HideSelection = false;
            this.treeViewProjectList.Location = new System.Drawing.Point(0, 0);
            this.treeViewProjectList.Name = "treeViewProjectList";
            this.treeViewProjectList.Size = new System.Drawing.Size(239, 480);
            this.treeViewProjectList.TabIndex = 0;
            // 
            // panelLeftTop
            // 
            this.panelLeftTop.AutoSize = true;
            this.panelLeftTop.Controls.Add(this.toolStripProjectList);
            this.panelLeftTop.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelLeftTop.Location = new System.Drawing.Point(1, 1);
            this.panelLeftTop.Name = "panelLeftTop";
            this.panelLeftTop.Size = new System.Drawing.Size(239, 25);
            this.panelLeftTop.TabIndex = 0;
            // 
            // toolStripProjectList
            // 
            this.toolStripProjectList.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripDropDownButtonAdd,
            this.toolStripButtonEdit,
            this.toolStripButtonDelete,
            this.toolStripSeparator1,
            this.toolStripButtonRun});
            this.toolStripProjectList.Location = new System.Drawing.Point(0, 0);
            this.toolStripProjectList.Name = "toolStripProjectList";
            this.toolStripProjectList.RenderMode = System.Windows.Forms.ToolStripRenderMode.System;
            this.toolStripProjectList.Size = new System.Drawing.Size(239, 25);
            this.toolStripProjectList.TabIndex = 0;
            this.toolStripProjectList.Text = "toolStripProjectList";
            // 
            // toolStripDropDownButtonAdd
            // 
            this.toolStripDropDownButtonAdd.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.toolStripDropDownButtonAdd.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItemAddGroup,
            this.toolStripMenuItemAddProject});
            this.toolStripDropDownButtonAdd.Image = ((System.Drawing.Image)(resources.GetObject("toolStripDropDownButtonAdd.Image")));
            this.toolStripDropDownButtonAdd.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripDropDownButtonAdd.Name = "toolStripDropDownButtonAdd";
            this.toolStripDropDownButtonAdd.Size = new System.Drawing.Size(63, 22);
            this.toolStripDropDownButtonAdd.Text = "新建(&N)";
            // 
            // toolStripMenuItemAddGroup
            // 
            this.toolStripMenuItemAddGroup.Name = "toolStripMenuItemAddGroup";
            this.toolStripMenuItemAddGroup.Size = new System.Drawing.Size(126, 22);
            this.toolStripMenuItemAddGroup.Text = "分组(&G)...";
            this.toolStripMenuItemAddGroup.Click += new System.EventHandler(this.toolStripMenuItemAddGroup_Click);
            // 
            // toolStripMenuItemAddProject
            // 
            this.toolStripMenuItemAddProject.Name = "toolStripMenuItemAddProject";
            this.toolStripMenuItemAddProject.Size = new System.Drawing.Size(126, 22);
            this.toolStripMenuItemAddProject.Text = "项目(&P)...";
            this.toolStripMenuItemAddProject.Click += new System.EventHandler(this.toolStripMenuItemAddProject_Click);
            // 
            // toolStripButtonEdit
            // 
            this.toolStripButtonEdit.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.toolStripButtonEdit.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButtonEdit.Image")));
            this.toolStripButtonEdit.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButtonEdit.Name = "toolStripButtonEdit";
            this.toolStripButtonEdit.Size = new System.Drawing.Size(51, 22);
            this.toolStripButtonEdit.Text = "编辑(&E)";
            this.toolStripButtonEdit.Click += new System.EventHandler(this.toolStripButtonEdit_Click);
            // 
            // toolStripButtonDelete
            // 
            this.toolStripButtonDelete.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.toolStripButtonDelete.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButtonDelete.Image")));
            this.toolStripButtonDelete.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButtonDelete.Name = "toolStripButtonDelete";
            this.toolStripButtonDelete.Size = new System.Drawing.Size(53, 22);
            this.toolStripButtonDelete.Text = "删除(&D)";
            this.toolStripButtonDelete.Click += new System.EventHandler(this.toolStripButtonDelete_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 25);
            // 
            // toolStripButtonRun
            // 
            this.toolStripButtonRun.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.toolStripButtonRun.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButtonRun.Image")));
            this.toolStripButtonRun.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButtonRun.Name = "toolStripButtonRun";
            this.toolStripButtonRun.Size = new System.Drawing.Size(52, 22);
            this.toolStripButtonRun.Text = "运行(&R)";
            this.toolStripButtonRun.Click += new System.EventHandler(this.toolStripButtonRun_Click);
            // 
            // tabControlMain
            // 
            this.tabControlMain.Controls.Add(this.tabPageMain);
            this.tabControlMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControlMain.Location = new System.Drawing.Point(0, 0);
            this.tabControlMain.Name = "tabControlMain";
            this.tabControlMain.SelectedIndex = 0;
            this.tabControlMain.Size = new System.Drawing.Size(767, 507);
            this.tabControlMain.TabIndex = 0;
            // 
            // tabPageMain
            // 
            this.tabPageMain.Location = new System.Drawing.Point(4, 22);
            this.tabPageMain.Name = "tabPageMain";
            this.tabPageMain.Size = new System.Drawing.Size(759, 481);
            this.tabPageMain.TabIndex = 0;
            this.tabPageMain.Text = "Home Page";
            this.tabPageMain.UseVisualStyleBackColor = true;
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.settingToolStripMenuItem,
            this.helpToolStripMenuItem,
            this.toolsToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1012, 25);
            this.menuStrip1.TabIndex = 2;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // settingToolStripMenuItem
            // 
            this.settingToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.configToolStripMenuItem});
            this.settingToolStripMenuItem.Name = "settingToolStripMenuItem";
            this.settingToolStripMenuItem.Size = new System.Drawing.Size(59, 21);
            this.settingToolStripMenuItem.Text = "系统(&S)";
            // 
            // configToolStripMenuItem
            // 
            this.configToolStripMenuItem.Name = "configToolStripMenuItem";
            this.configToolStripMenuItem.Size = new System.Drawing.Size(125, 22);
            this.configToolStripMenuItem.Text = "配置(&C)...";
            this.configToolStripMenuItem.Click += new System.EventHandler(this.configToolStripMenuItem_Click);
            // 
            // helpToolStripMenuItem
            // 
            this.helpToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.aboutToolStripMenuItem});
            this.helpToolStripMenuItem.Name = "helpToolStripMenuItem";
            this.helpToolStripMenuItem.Size = new System.Drawing.Size(61, 21);
            this.helpToolStripMenuItem.Text = "帮助(&H)";
            // 
            // aboutToolStripMenuItem
            // 
            this.aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
            this.aboutToolStripMenuItem.Size = new System.Drawing.Size(116, 22);
            this.aboutToolStripMenuItem.Text = "关于(&A)";
            this.aboutToolStripMenuItem.Click += new System.EventHandler(this.aboutToolStripMenuItem_Click);
            // 
            // toolsToolStripMenuItem
            // 
            this.toolsToolStripMenuItem.Name = "toolsToolStripMenuItem";
            this.toolsToolStripMenuItem.Size = new System.Drawing.Size(59, 21);
            this.toolsToolStripMenuItem.Text = "工具(&T)";
            // 
            // timerServer
            // 
            this.timerServer.Enabled = true;
            this.timerServer.Interval = 5000;
            this.timerServer.Tick += new System.EventHandler(this.timerServer_Tick);
            // 
            // statusStripMain
            // 
            this.statusStripMain.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabelServerStatus});
            this.statusStripMain.Location = new System.Drawing.Point(0, 532);
            this.statusStripMain.Name = "statusStripMain";
            this.statusStripMain.Size = new System.Drawing.Size(1012, 22);
            this.statusStripMain.TabIndex = 3;
            this.statusStripMain.Text = "statusStrip1";
            // 
            // toolStripStatusLabelServerStatus
            // 
            this.toolStripStatusLabelServerStatus.Name = "toolStripStatusLabelServerStatus";
            this.toolStripStatusLabelServerStatus.Size = new System.Drawing.Size(80, 17);
            this.toolStripStatusLabelServerStatus.Text = "服务接口状态";
            // 
            // FormMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1012, 554);
            this.Controls.Add(this.splitContainerMain);
            this.Controls.Add(this.menuStrip1);
            this.Controls.Add(this.statusStripMain);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "FormMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Net Data Access";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.splitContainerMain.Panel1.ResumeLayout(false);
            this.splitContainerMain.Panel1.PerformLayout();
            this.splitContainerMain.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainerMain)).EndInit();
            this.splitContainerMain.ResumeLayout(false);
            this.panelLeftCenter.ResumeLayout(false);
            this.panelLeftTop.ResumeLayout(false);
            this.panelLeftTop.PerformLayout();
            this.toolStripProjectList.ResumeLayout(false);
            this.toolStripProjectList.PerformLayout();
            this.tabControlMain.ResumeLayout(false);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.statusStripMain.ResumeLayout(false);
            this.statusStripMain.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.SplitContainer splitContainerMain;
        private System.Windows.Forms.Panel panelLeftTop;
        private System.Windows.Forms.Panel panelLeftCenter;
        private System.Windows.Forms.TabControl tabControlMain;
        private System.Windows.Forms.TabPage tabPageMain;
        private System.Windows.Forms.ToolStrip toolStripProjectList;
        private System.Windows.Forms.ToolStripButton toolStripButtonDelete;
        private System.Windows.Forms.ToolStripButton toolStripButtonRun;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.TreeView treeViewProjectList;
        private System.Windows.Forms.ToolStripDropDownButton toolStripDropDownButtonAdd;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItemAddGroup;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItemAddProject;
        private System.Windows.Forms.ToolStripButton toolStripButtonEdit;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem settingToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem configToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem helpToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem aboutToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem toolsToolStripMenuItem;
        private System.Windows.Forms.Timer timerServer;
        private System.Windows.Forms.StatusStrip statusStripMain;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabelServerStatus;
    }
}

