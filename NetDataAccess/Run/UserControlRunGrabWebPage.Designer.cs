namespace NetDataAccess.Run
{
    partial class UserControlRunGrabWebPage
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

        #region 组件设计器生成的代码

        /// <summary> 
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.tabPageGrabLog = new System.Windows.Forms.TabPage();
            this.textBoxGrabLog = new System.Windows.Forms.TextBox();
            this.textBoxStatus = new System.Windows.Forms.TextBox();
            this.panelMain = new System.Windows.Forms.Panel();
            this.tabControlMain = new System.Windows.Forms.TabControl();
            this.tabPageGrabLog.SuspendLayout();
            this.panelMain.SuspendLayout();
            this.tabControlMain.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabPageGrabLog
            // 
            this.tabPageGrabLog.Controls.Add(this.textBoxGrabLog);
            this.tabPageGrabLog.Controls.Add(this.textBoxStatus);
            this.tabPageGrabLog.Location = new System.Drawing.Point(4, 22);
            this.tabPageGrabLog.Name = "tabPageGrabLog";
            this.tabPageGrabLog.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageGrabLog.Size = new System.Drawing.Size(947, 531);
            this.tabPageGrabLog.TabIndex = 1;
            this.tabPageGrabLog.Text = "进度";
            this.tabPageGrabLog.UseVisualStyleBackColor = true;
            // 
            // textBoxGrabLog
            // 
            this.textBoxGrabLog.Dock = System.Windows.Forms.DockStyle.Fill;
            this.textBoxGrabLog.Font = new System.Drawing.Font("宋体", 9F);
            this.textBoxGrabLog.Location = new System.Drawing.Point(3, 24);
            this.textBoxGrabLog.MaxLength = 3276700;
            this.textBoxGrabLog.Multiline = true;
            this.textBoxGrabLog.Name = "textBoxGrabLog";
            this.textBoxGrabLog.ReadOnly = true;
            this.textBoxGrabLog.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBoxGrabLog.Size = new System.Drawing.Size(941, 504);
            this.textBoxGrabLog.TabIndex = 1;
            // 
            // textBoxStatus
            // 
            this.textBoxStatus.Dock = System.Windows.Forms.DockStyle.Top;
            this.textBoxStatus.Font = new System.Drawing.Font("宋体", 9F);
            this.textBoxStatus.Location = new System.Drawing.Point(3, 3);
            this.textBoxStatus.MaxLength = 3276700;
            this.textBoxStatus.Name = "textBoxStatus";
            this.textBoxStatus.ReadOnly = true;
            this.textBoxStatus.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBoxStatus.Size = new System.Drawing.Size(941, 21);
            this.textBoxStatus.TabIndex = 2;
            // 
            // panelMain
            // 
            this.panelMain.Controls.Add(this.tabControlMain);
            this.panelMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelMain.Location = new System.Drawing.Point(0, 0);
            this.panelMain.Name = "panelMain";
            this.panelMain.Size = new System.Drawing.Size(955, 557);
            this.panelMain.TabIndex = 3;
            // 
            // tabControlMain
            // 
            this.tabControlMain.Controls.Add(this.tabPageGrabLog);
            this.tabControlMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControlMain.Location = new System.Drawing.Point(0, 0);
            this.tabControlMain.Name = "tabControlMain";
            this.tabControlMain.SelectedIndex = 0;
            this.tabControlMain.Size = new System.Drawing.Size(955, 557);
            this.tabControlMain.TabIndex = 1;
            // 
            // UserControlRunGrabWebPage
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.panelMain);
            this.Name = "UserControlRunGrabWebPage";
            this.Size = new System.Drawing.Size(955, 557);
            this.tabPageGrabLog.ResumeLayout(false);
            this.tabPageGrabLog.PerformLayout();
            this.panelMain.ResumeLayout(false);
            this.tabControlMain.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabPage tabPageGrabLog;
        private System.Windows.Forms.Panel panelMain;
        private System.Windows.Forms.TabControl tabControlMain;
        private System.Windows.Forms.TextBox textBoxGrabLog;
        private System.Windows.Forms.TextBox textBoxStatus;
    }
}
