namespace PictureToPDF
{
    partial class MainForm
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置 Managed 資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器
        /// 修改這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.labelX1 = new DevComponents.DotNetBar.LabelX();
            this.txtSourcePath = new System.Windows.Forms.TextBox();
            this.btnSource = new DevComponents.DotNetBar.ButtonX();
            this.btnTarget = new DevComponents.DotNetBar.ButtonX();
            this.txtTargetPath = new System.Windows.Forms.TextBox();
            this.labelX2 = new DevComponents.DotNetBar.LabelX();
            this.btnConfirm = new DevComponents.DotNetBar.ButtonX();
            this.btnExit = new DevComponents.DotNetBar.ButtonX();
            this.btnPicFolder = new DevComponents.DotNetBar.ButtonX();
            this.txtPicFolder = new System.Windows.Forms.TextBox();
            this.labelX3 = new DevComponents.DotNetBar.LabelX();
            this.SuspendLayout();
            // 
            // labelX1
            // 
            this.labelX1.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX1.BackgroundStyle.Class = "";
            this.labelX1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX1.Location = new System.Drawing.Point(13, 13);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(341, 23);
            this.labelX1.TabIndex = 8;
            this.labelX1.Text = "請選擇設定檔案路徑(*.xls)";
            // 
            // txtSourcePath
            // 
            this.txtSourcePath.Location = new System.Drawing.Point(13, 43);
            this.txtSourcePath.Name = "txtSourcePath";
            this.txtSourcePath.Size = new System.Drawing.Size(435, 25);
            this.txtSourcePath.TabIndex = 1;
            this.txtSourcePath.TextChanged += new System.EventHandler(this.txtSourcePath_TextChanged);
            // 
            // btnSource
            // 
            this.btnSource.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnSource.BackColor = System.Drawing.Color.Transparent;
            this.btnSource.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnSource.Location = new System.Drawing.Point(373, 13);
            this.btnSource.Name = "btnSource";
            this.btnSource.Size = new System.Drawing.Size(75, 23);
            this.btnSource.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnSource.TabIndex = 0;
            this.btnSource.Text = "選擇";
            this.btnSource.Click += new System.EventHandler(this.btnSource_Click);
            // 
            // btnTarget
            // 
            this.btnTarget.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnTarget.BackColor = System.Drawing.Color.Transparent;
            this.btnTarget.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnTarget.Location = new System.Drawing.Point(373, 135);
            this.btnTarget.Name = "btnTarget";
            this.btnTarget.Size = new System.Drawing.Size(75, 23);
            this.btnTarget.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnTarget.TabIndex = 2;
            this.btnTarget.Text = "選擇";
            this.btnTarget.Click += new System.EventHandler(this.btnTarget_Click);
            // 
            // txtTargetPath
            // 
            this.txtTargetPath.Location = new System.Drawing.Point(13, 165);
            this.txtTargetPath.Name = "txtTargetPath";
            this.txtTargetPath.Size = new System.Drawing.Size(435, 25);
            this.txtTargetPath.TabIndex = 3;
            this.txtTargetPath.TextChanged += new System.EventHandler(this.txtTargetPath_TextChanged);
            // 
            // labelX2
            // 
            this.labelX2.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX2.BackgroundStyle.Class = "";
            this.labelX2.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX2.Location = new System.Drawing.Point(13, 135);
            this.labelX2.Name = "labelX2";
            this.labelX2.Size = new System.Drawing.Size(341, 23);
            this.labelX2.TabIndex = 9;
            this.labelX2.Text = "請選擇輸出目的資料夾";
            // 
            // btnConfirm
            // 
            this.btnConfirm.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnConfirm.BackColor = System.Drawing.Color.Transparent;
            this.btnConfirm.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnConfirm.Location = new System.Drawing.Point(290, 210);
            this.btnConfirm.Name = "btnConfirm";
            this.btnConfirm.Size = new System.Drawing.Size(75, 23);
            this.btnConfirm.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnConfirm.TabIndex = 6;
            this.btnConfirm.Text = "確認";
            this.btnConfirm.Click += new System.EventHandler(this.btnConfirm_Click);
            // 
            // btnExit
            // 
            this.btnExit.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnExit.BackColor = System.Drawing.Color.Transparent;
            this.btnExit.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnExit.Location = new System.Drawing.Point(373, 210);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(75, 23);
            this.btnExit.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnExit.TabIndex = 7;
            this.btnExit.Text = "離開";
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // btnPicFolder
            // 
            this.btnPicFolder.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnPicFolder.BackColor = System.Drawing.Color.Transparent;
            this.btnPicFolder.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnPicFolder.Location = new System.Drawing.Point(372, 74);
            this.btnPicFolder.Name = "btnPicFolder";
            this.btnPicFolder.Size = new System.Drawing.Size(75, 23);
            this.btnPicFolder.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnPicFolder.TabIndex = 4;
            this.btnPicFolder.Text = "選擇";
            this.btnPicFolder.Click += new System.EventHandler(this.btnPicFolder_Click);
            // 
            // txtPicFolder
            // 
            this.txtPicFolder.Location = new System.Drawing.Point(12, 104);
            this.txtPicFolder.Name = "txtPicFolder";
            this.txtPicFolder.Size = new System.Drawing.Size(435, 25);
            this.txtPicFolder.TabIndex = 5;
            this.txtPicFolder.TextChanged += new System.EventHandler(this.txtPicFolder_TextChanged);
            // 
            // labelX3
            // 
            this.labelX3.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX3.BackgroundStyle.Class = "";
            this.labelX3.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX3.Location = new System.Drawing.Point(12, 74);
            this.labelX3.Name = "labelX3";
            this.labelX3.Size = new System.Drawing.Size(341, 23);
            this.labelX3.TabIndex = 10;
            this.labelX3.Text = "請選擇圖片來源資料夾";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(461, 240);
            this.Controls.Add(this.btnPicFolder);
            this.Controls.Add(this.txtPicFolder);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnConfirm);
            this.Controls.Add(this.btnTarget);
            this.Controls.Add(this.txtTargetPath);
            this.Controls.Add(this.btnSource);
            this.Controls.Add(this.txtSourcePath);
            this.Controls.Add(this.labelX3);
            this.Controls.Add(this.labelX2);
            this.Controls.Add(this.labelX1);
            this.DoubleBuffered = true;
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "設定";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevComponents.DotNetBar.LabelX labelX1;
        private System.Windows.Forms.TextBox txtSourcePath;
        private DevComponents.DotNetBar.ButtonX btnSource;
        private DevComponents.DotNetBar.ButtonX btnTarget;
        private System.Windows.Forms.TextBox txtTargetPath;
        private DevComponents.DotNetBar.LabelX labelX2;
        private DevComponents.DotNetBar.ButtonX btnConfirm;
        private DevComponents.DotNetBar.ButtonX btnExit;
        private DevComponents.DotNetBar.ButtonX btnPicFolder;
        private System.Windows.Forms.TextBox txtPicFolder;
        private DevComponents.DotNetBar.LabelX labelX3;
    }
}

