namespace NetDataAccess.Edit
{
    partial class UserControlEditProject
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
            this.panelLoginTop = new System.Windows.Forms.Panel();
            this.comboBoxLoginType = new System.Windows.Forms.ComboBox();
            this.labelLoginType = new System.Windows.Forms.Label();
            this.labelLoginPageInfo = new System.Windows.Forms.Label();
            this.groupBoxBase = new System.Windows.Forms.GroupBox();
            this.textBoxDescription = new System.Windows.Forms.TextBox();
            this.labelDescription = new System.Windows.Forms.Label();
            this.textBoxName = new System.Windows.Forms.TextBox();
            this.labelName = new System.Windows.Forms.Label();
            this.panelBottom = new System.Windows.Forms.Panel();
            this.buttonClose = new System.Windows.Forms.Button();
            this.buttonOK = new System.Windows.Forms.Button();
            this.textBoxLoginPageInfo = new System.Windows.Forms.TextBox();
            this.panelMain = new System.Windows.Forms.Panel();
            this.groupBoxExternal = new System.Windows.Forms.GroupBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.textBoxProgramExternalRun = new System.Windows.Forms.TextBox();
            this.labelProgramExternalRun = new System.Windows.Forms.Label();
            this.labelProgramAfterGrabAll = new System.Windows.Forms.Label();
            this.textBoxProgramAfterGrabAll = new System.Windows.Forms.TextBox();
            this.groupBoxDetailPage = new System.Windows.Forms.GroupBox();
            this.panelDetailBottom = new System.Windows.Forms.Panel();
            this.labelDetailGrabInfo = new System.Windows.Forms.Label();
            this.textBoxDetailGrabInfo = new System.Windows.Forms.TextBox();
            this.panelDetailTop = new System.Windows.Forms.Panel();
            this.comboBoxDetailGrabType = new System.Windows.Forms.ComboBox();
            this.labelDetailGrabType = new System.Windows.Forms.Label();
            this.groupBoxLogin = new System.Windows.Forms.GroupBox();
            this.panelLoginBottom = new System.Windows.Forms.Panel();
            this.panelLoginTop.SuspendLayout();
            this.groupBoxBase.SuspendLayout();
            this.panelBottom.SuspendLayout();
            this.panelMain.SuspendLayout();
            this.groupBoxExternal.SuspendLayout();
            this.panel1.SuspendLayout();
            this.groupBoxDetailPage.SuspendLayout();
            this.panelDetailBottom.SuspendLayout();
            this.panelDetailTop.SuspendLayout();
            this.groupBoxLogin.SuspendLayout();
            this.panelLoginBottom.SuspendLayout();
            this.SuspendLayout();
            // 
            // panelLoginTop
            // 
            this.panelLoginTop.Controls.Add(this.comboBoxLoginType);
            this.panelLoginTop.Controls.Add(this.labelLoginType);
            this.panelLoginTop.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelLoginTop.Location = new System.Drawing.Point(3, 17);
            this.panelLoginTop.Name = "panelLoginTop";
            this.panelLoginTop.Size = new System.Drawing.Size(933, 23);
            this.panelLoginTop.TabIndex = 5;
            // 
            // comboBoxLoginType
            // 
            this.comboBoxLoginType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxLoginType.FormattingEnabled = true;
            this.comboBoxLoginType.Items.AddRange(new object[] {
            "HasLogin",
            "NoneLogin"});
            this.comboBoxLoginType.Location = new System.Drawing.Point(105, 2);
            this.comboBoxLoginType.Name = "comboBoxLoginType";
            this.comboBoxLoginType.Size = new System.Drawing.Size(163, 20);
            this.comboBoxLoginType.TabIndex = 6;
            // 
            // labelLoginType
            // 
            this.labelLoginType.AutoSize = true;
            this.labelLoginType.Location = new System.Drawing.Point(43, 7);
            this.labelLoginType.Name = "labelLoginType";
            this.labelLoginType.Size = new System.Drawing.Size(59, 12);
            this.labelLoginType.TabIndex = 3;
            this.labelLoginType.Text = "登录方式:";
            // 
            // labelLoginPageInfo
            // 
            this.labelLoginPageInfo.AutoSize = true;
            this.labelLoginPageInfo.Location = new System.Drawing.Point(33, 6);
            this.labelLoginPageInfo.Name = "labelLoginPageInfo";
            this.labelLoginPageInfo.Size = new System.Drawing.Size(71, 12);
            this.labelLoginPageInfo.TabIndex = 1;
            this.labelLoginPageInfo.Text = "登录页设置:";
            // 
            // groupBoxBase
            // 
            this.groupBoxBase.Controls.Add(this.textBoxDescription);
            this.groupBoxBase.Controls.Add(this.labelDescription);
            this.groupBoxBase.Controls.Add(this.textBoxName);
            this.groupBoxBase.Controls.Add(this.labelName);
            this.groupBoxBase.Dock = System.Windows.Forms.DockStyle.Top;
            this.groupBoxBase.Location = new System.Drawing.Point(0, 0);
            this.groupBoxBase.Name = "groupBoxBase";
            this.groupBoxBase.Size = new System.Drawing.Size(939, 100);
            this.groupBoxBase.TabIndex = 3;
            this.groupBoxBase.TabStop = false;
            this.groupBoxBase.Text = "基本信息";
            // 
            // textBoxDescription
            // 
            this.textBoxDescription.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxDescription.Location = new System.Drawing.Point(108, 45);
            this.textBoxDescription.Multiline = true;
            this.textBoxDescription.Name = "textBoxDescription";
            this.textBoxDescription.Size = new System.Drawing.Size(818, 49);
            this.textBoxDescription.TabIndex = 1;
            // 
            // labelDescription
            // 
            this.labelDescription.AutoSize = true;
            this.labelDescription.Location = new System.Drawing.Point(70, 48);
            this.labelDescription.Name = "labelDescription";
            this.labelDescription.Size = new System.Drawing.Size(35, 12);
            this.labelDescription.TabIndex = 1;
            this.labelDescription.Text = "说明:";
            // 
            // textBoxName
            // 
            this.textBoxName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxName.Location = new System.Drawing.Point(108, 18);
            this.textBoxName.Name = "textBoxName";
            this.textBoxName.Size = new System.Drawing.Size(818, 21);
            this.textBoxName.TabIndex = 0;
            // 
            // labelName
            // 
            this.labelName.AutoSize = true;
            this.labelName.Location = new System.Drawing.Point(70, 21);
            this.labelName.Name = "labelName";
            this.labelName.Size = new System.Drawing.Size(35, 12);
            this.labelName.TabIndex = 1;
            this.labelName.Text = "名称:";
            // 
            // panelBottom
            // 
            this.panelBottom.Controls.Add(this.buttonClose);
            this.panelBottom.Controls.Add(this.buttonOK);
            this.panelBottom.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelBottom.Location = new System.Drawing.Point(0, 0);
            this.panelBottom.Name = "panelBottom";
            this.panelBottom.Size = new System.Drawing.Size(939, 32);
            this.panelBottom.TabIndex = 3;
            // 
            // buttonClose
            // 
            this.buttonClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonClose.Location = new System.Drawing.Point(861, 3);
            this.buttonClose.Name = "buttonClose";
            this.buttonClose.Size = new System.Drawing.Size(75, 23);
            this.buttonClose.TabIndex = 13;
            this.buttonClose.Text = "关闭(&C)";
            this.buttonClose.UseVisualStyleBackColor = true;
            this.buttonClose.Click += new System.EventHandler(this.buttonClose_Click);
            // 
            // buttonOK
            // 
            this.buttonOK.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonOK.Location = new System.Drawing.Point(780, 3);
            this.buttonOK.Name = "buttonOK";
            this.buttonOK.Size = new System.Drawing.Size(75, 23);
            this.buttonOK.TabIndex = 12;
            this.buttonOK.Text = "保存(&S)";
            this.buttonOK.UseVisualStyleBackColor = true;
            this.buttonOK.Click += new System.EventHandler(this.buttonOK_Click);
            // 
            // textBoxLoginPageInfo
            // 
            this.textBoxLoginPageInfo.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxLoginPageInfo.Location = new System.Drawing.Point(105, 3);
            this.textBoxLoginPageInfo.Multiline = true;
            this.textBoxLoginPageInfo.Name = "textBoxLoginPageInfo";
            this.textBoxLoginPageInfo.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBoxLoginPageInfo.Size = new System.Drawing.Size(820, 49);
            this.textBoxLoginPageInfo.TabIndex = 7;
            // 
            // panelMain
            // 
            this.panelMain.Controls.Add(this.groupBoxExternal);
            this.panelMain.Controls.Add(this.groupBoxDetailPage);
            this.panelMain.Controls.Add(this.groupBoxLogin);
            this.panelMain.Controls.Add(this.groupBoxBase);
            this.panelMain.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelMain.Location = new System.Drawing.Point(0, 32);
            this.panelMain.Name = "panelMain";
            this.panelMain.Size = new System.Drawing.Size(939, 511);
            this.panelMain.TabIndex = 2;
            // 
            // groupBoxExternal
            // 
            this.groupBoxExternal.AutoSize = true;
            this.groupBoxExternal.Controls.Add(this.panel1);
            this.groupBoxExternal.Dock = System.Windows.Forms.DockStyle.Top;
            this.groupBoxExternal.Location = new System.Drawing.Point(0, 297);
            this.groupBoxExternal.Name = "groupBoxExternal";
            this.groupBoxExternal.Size = new System.Drawing.Size(939, 133);
            this.groupBoxExternal.TabIndex = 9;
            this.groupBoxExternal.TabStop = false;
            this.groupBoxExternal.Text = "扩展程序";
            // 
            // panel1
            // 
            this.panel1.AutoSize = true;
            this.panel1.Controls.Add(this.textBoxProgramExternalRun);
            this.panel1.Controls.Add(this.labelProgramExternalRun);
            this.panel1.Controls.Add(this.labelProgramAfterGrabAll);
            this.panel1.Controls.Add(this.textBoxProgramAfterGrabAll);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(3, 17);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(933, 113);
            this.panel1.TabIndex = 5;
            // 
            // textBoxProgramExternalRun
            // 
            this.textBoxProgramExternalRun.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxProgramExternalRun.Location = new System.Drawing.Point(105, 6);
            this.textBoxProgramExternalRun.Multiline = true;
            this.textBoxProgramExternalRun.Name = "textBoxProgramExternalRun";
            this.textBoxProgramExternalRun.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBoxProgramExternalRun.Size = new System.Drawing.Size(818, 49);
            this.textBoxProgramExternalRun.TabIndex = 12;
            // 
            // labelProgramExternalRun
            // 
            this.labelProgramExternalRun.AutoSize = true;
            this.labelProgramExternalRun.Location = new System.Drawing.Point(45, 6);
            this.labelProgramExternalRun.Name = "labelProgramExternalRun";
            this.labelProgramExternalRun.Size = new System.Drawing.Size(59, 12);
            this.labelProgramExternalRun.TabIndex = 1;
            this.labelProgramExternalRun.Text = "扩展设置:";
            // 
            // labelProgramAfterGrabAll
            // 
            this.labelProgramAfterGrabAll.Location = new System.Drawing.Point(3, 61);
            this.labelProgramAfterGrabAll.Name = "labelProgramAfterGrabAll";
            this.labelProgramAfterGrabAll.Size = new System.Drawing.Size(101, 32);
            this.labelProgramAfterGrabAll.TabIndex = 1;
            this.labelProgramAfterGrabAll.Text = "全部抓取完成后:\r\n(即将作废)";
            this.labelProgramAfterGrabAll.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // textBoxProgramAfterGrabAll
            // 
            this.textBoxProgramAfterGrabAll.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxProgramAfterGrabAll.Location = new System.Drawing.Point(105, 61);
            this.textBoxProgramAfterGrabAll.Multiline = true;
            this.textBoxProgramAfterGrabAll.Name = "textBoxProgramAfterGrabAll";
            this.textBoxProgramAfterGrabAll.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBoxProgramAfterGrabAll.Size = new System.Drawing.Size(816, 49);
            this.textBoxProgramAfterGrabAll.TabIndex = 11;
            // 
            // groupBoxDetailPage
            // 
            this.groupBoxDetailPage.AutoSize = true;
            this.groupBoxDetailPage.Controls.Add(this.panelDetailBottom);
            this.groupBoxDetailPage.Controls.Add(this.panelDetailTop);
            this.groupBoxDetailPage.Dock = System.Windows.Forms.DockStyle.Top;
            this.groupBoxDetailPage.Location = new System.Drawing.Point(0, 198);
            this.groupBoxDetailPage.Name = "groupBoxDetailPage";
            this.groupBoxDetailPage.Size = new System.Drawing.Size(939, 99);
            this.groupBoxDetailPage.TabIndex = 8;
            this.groupBoxDetailPage.TabStop = false;
            this.groupBoxDetailPage.Text = "详情页";
            // 
            // panelDetailBottom
            // 
            this.panelDetailBottom.AutoSize = true;
            this.panelDetailBottom.Controls.Add(this.labelDetailGrabInfo);
            this.panelDetailBottom.Controls.Add(this.textBoxDetailGrabInfo);
            this.panelDetailBottom.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelDetailBottom.Location = new System.Drawing.Point(3, 41);
            this.panelDetailBottom.Name = "panelDetailBottom";
            this.panelDetailBottom.Size = new System.Drawing.Size(933, 55);
            this.panelDetailBottom.TabIndex = 5;
            // 
            // labelDetailGrabInfo
            // 
            this.labelDetailGrabInfo.AutoSize = true;
            this.labelDetailGrabInfo.Location = new System.Drawing.Point(45, 6);
            this.labelDetailGrabInfo.Name = "labelDetailGrabInfo";
            this.labelDetailGrabInfo.Size = new System.Drawing.Size(59, 12);
            this.labelDetailGrabInfo.TabIndex = 1;
            this.labelDetailGrabInfo.Text = "抓取设置:";
            // 
            // textBoxDetailGrabInfo
            // 
            this.textBoxDetailGrabInfo.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxDetailGrabInfo.Location = new System.Drawing.Point(107, 3);
            this.textBoxDetailGrabInfo.Multiline = true;
            this.textBoxDetailGrabInfo.Name = "textBoxDetailGrabInfo";
            this.textBoxDetailGrabInfo.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBoxDetailGrabInfo.Size = new System.Drawing.Size(816, 49);
            this.textBoxDetailGrabInfo.TabIndex = 11;
            // 
            // panelDetailTop
            // 
            this.panelDetailTop.Controls.Add(this.comboBoxDetailGrabType);
            this.panelDetailTop.Controls.Add(this.labelDetailGrabType);
            this.panelDetailTop.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelDetailTop.Location = new System.Drawing.Point(3, 17);
            this.panelDetailTop.Name = "panelDetailTop";
            this.panelDetailTop.Size = new System.Drawing.Size(933, 24);
            this.panelDetailTop.TabIndex = 5;
            // 
            // comboBoxDetailGrabType
            // 
            this.comboBoxDetailGrabType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxDetailGrabType.FormattingEnabled = true;
            this.comboBoxDetailGrabType.Items.AddRange(new object[] {
            "SingleLineType",
            "MultiLineType",
            "NoneDetailPage",
            "ProgramType"});
            this.comboBoxDetailGrabType.Location = new System.Drawing.Point(107, 3);
            this.comboBoxDetailGrabType.Name = "comboBoxDetailGrabType";
            this.comboBoxDetailGrabType.Size = new System.Drawing.Size(163, 20);
            this.comboBoxDetailGrabType.TabIndex = 10;
            // 
            // labelDetailGrabType
            // 
            this.labelDetailGrabType.AutoSize = true;
            this.labelDetailGrabType.Location = new System.Drawing.Point(45, 6);
            this.labelDetailGrabType.Name = "labelDetailGrabType";
            this.labelDetailGrabType.Size = new System.Drawing.Size(59, 12);
            this.labelDetailGrabType.TabIndex = 3;
            this.labelDetailGrabType.Text = "抓取方式:";
            // 
            // groupBoxLogin
            // 
            this.groupBoxLogin.AutoSize = true;
            this.groupBoxLogin.Controls.Add(this.panelLoginBottom);
            this.groupBoxLogin.Controls.Add(this.panelLoginTop);
            this.groupBoxLogin.Dock = System.Windows.Forms.DockStyle.Top;
            this.groupBoxLogin.Location = new System.Drawing.Point(0, 100);
            this.groupBoxLogin.Name = "groupBoxLogin";
            this.groupBoxLogin.Size = new System.Drawing.Size(939, 98);
            this.groupBoxLogin.TabIndex = 4;
            this.groupBoxLogin.TabStop = false;
            this.groupBoxLogin.Text = "登录";
            // 
            // panelLoginBottom
            // 
            this.panelLoginBottom.AutoSize = true;
            this.panelLoginBottom.Controls.Add(this.labelLoginPageInfo);
            this.panelLoginBottom.Controls.Add(this.textBoxLoginPageInfo);
            this.panelLoginBottom.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelLoginBottom.Location = new System.Drawing.Point(3, 40);
            this.panelLoginBottom.Name = "panelLoginBottom";
            this.panelLoginBottom.Size = new System.Drawing.Size(933, 55);
            this.panelLoginBottom.TabIndex = 5;
            // 
            // UserControlEditProject
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.panelMain);
            this.Controls.Add(this.panelBottom);
            this.Name = "UserControlEditProject";
            this.Size = new System.Drawing.Size(939, 572);
            this.panelLoginTop.ResumeLayout(false);
            this.panelLoginTop.PerformLayout();
            this.groupBoxBase.ResumeLayout(false);
            this.groupBoxBase.PerformLayout();
            this.panelBottom.ResumeLayout(false);
            this.panelMain.ResumeLayout(false);
            this.panelMain.PerformLayout();
            this.groupBoxExternal.ResumeLayout(false);
            this.groupBoxExternal.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.groupBoxDetailPage.ResumeLayout(false);
            this.groupBoxDetailPage.PerformLayout();
            this.panelDetailBottom.ResumeLayout(false);
            this.panelDetailBottom.PerformLayout();
            this.panelDetailTop.ResumeLayout(false);
            this.panelDetailTop.PerformLayout();
            this.groupBoxLogin.ResumeLayout(false);
            this.groupBoxLogin.PerformLayout();
            this.panelLoginBottom.ResumeLayout(false);
            this.panelLoginBottom.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panelLoginTop;
        private System.Windows.Forms.ComboBox comboBoxLoginType;
        private System.Windows.Forms.Label labelLoginType;
        private System.Windows.Forms.Label labelLoginPageInfo;
        private System.Windows.Forms.GroupBox groupBoxBase;
        private System.Windows.Forms.TextBox textBoxDescription;
        private System.Windows.Forms.Label labelDescription;
        private System.Windows.Forms.TextBox textBoxName;
        private System.Windows.Forms.Label labelName;
        private System.Windows.Forms.Panel panelBottom;
        private System.Windows.Forms.Button buttonOK;
        private System.Windows.Forms.TextBox textBoxLoginPageInfo;
        private System.Windows.Forms.Panel panelMain;
        private System.Windows.Forms.GroupBox groupBoxExternal;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label labelProgramAfterGrabAll;
        private System.Windows.Forms.TextBox textBoxProgramAfterGrabAll;
        private System.Windows.Forms.GroupBox groupBoxDetailPage;
        private System.Windows.Forms.Panel panelDetailBottom;
        private System.Windows.Forms.Label labelDetailGrabInfo;
        private System.Windows.Forms.TextBox textBoxDetailGrabInfo;
        private System.Windows.Forms.Panel panelDetailTop;
        private System.Windows.Forms.ComboBox comboBoxDetailGrabType;
        private System.Windows.Forms.Label labelDetailGrabType;
        private System.Windows.Forms.GroupBox groupBoxLogin;
        private System.Windows.Forms.Panel panelLoginBottom;
        private System.Windows.Forms.Button buttonClose;
        private System.Windows.Forms.TextBox textBoxProgramExternalRun;
        private System.Windows.Forms.Label labelProgramExternalRun;

    }
}
