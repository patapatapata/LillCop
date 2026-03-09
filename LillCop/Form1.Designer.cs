namespace LillCop
{
    partial class Form1
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージ リソースが破棄される場合 true、破棄されない場合は false です。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows フォーム デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.button1 = new System.Windows.Forms.Button();
            this.buttonShowILL = new System.Windows.Forms.Button();
            this.Setting = new System.Windows.Forms.Button();
            this.ExpShow = new System.Windows.Forms.Button();
            this.Folder = new System.Windows.Forms.Button();
            this.button6 = new System.Windows.Forms.Button();
            this.show = new System.Windows.Forms.Button();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.label4 = new System.Windows.Forms.Label();
            this.buttonApp4 = new System.Windows.Forms.Button();
            this.buttonApp1 = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.button2 = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            this.ReLink = new System.Windows.Forms.Button();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.pictureSame = new System.Windows.Forms.PictureBox();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.buttonApp2 = new System.Windows.Forms.Button();
            this.buttonApp3 = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureSame)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            this.SuspendLayout();
            // 
            // timer1
            // 
            this.timer1.Interval = 500;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // backgroundWorker1
            // 
            this.backgroundWorker1.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker1_DoWork);
            // 
            // toolTip1
            // 
            this.toolTip1.AutomaticDelay = 250;
            this.toolTip1.AutoPopDelay = 10000;
            this.toolTip1.InitialDelay = 250;
            this.toolTip1.ReshowDelay = 50;
            // 
            // button1
            // 
            this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button1.FlatAppearance.BorderSize = 0;
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
            this.button1.Location = new System.Drawing.Point(276, 4);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(20, 19);
            this.button1.TabIndex = 41;
            this.toolTip1.SetToolTip(this.button1, "リンクファイルのリンク先階層をチェック");
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // buttonShowILL
            // 
            this.buttonShowILL.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonShowILL.FlatAppearance.BorderSize = 0;
            this.buttonShowILL.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonShowILL.Image = ((System.Drawing.Image)(resources.GetObject("buttonShowILL.Image")));
            this.buttonShowILL.Location = new System.Drawing.Point(301, 4);
            this.buttonShowILL.Name = "buttonShowILL";
            this.buttonShowILL.Size = new System.Drawing.Size(20, 19);
            this.buttonShowILL.TabIndex = 40;
            this.toolTip1.SetToolTip(this.buttonShowILL, "Illustratorドキュメントをエクスプローラーで表示");
            this.buttonShowILL.UseVisualStyleBackColor = true;
            this.buttonShowILL.Click += new System.EventHandler(this.button1_Click);
            // 
            // Setting
            // 
            this.Setting.FlatAppearance.BorderSize = 0;
            this.Setting.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.Setting.Image = ((System.Drawing.Image)(resources.GetObject("Setting.Image")));
            this.Setting.Location = new System.Drawing.Point(182, 4);
            this.Setting.Name = "Setting";
            this.Setting.Size = new System.Drawing.Size(23, 23);
            this.Setting.TabIndex = 39;
            this.toolTip1.SetToolTip(this.Setting, "設定");
            this.Setting.UseVisualStyleBackColor = true;
            this.Setting.Click += new System.EventHandler(this.Setting_Click);
            // 
            // ExpShow
            // 
            this.ExpShow.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.ExpShow.FlatAppearance.BorderSize = 0;
            this.ExpShow.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.ExpShow.Image = ((System.Drawing.Image)(resources.GetObject("ExpShow.Image")));
            this.ExpShow.Location = new System.Drawing.Point(301, 27);
            this.ExpShow.Name = "ExpShow";
            this.ExpShow.Size = new System.Drawing.Size(20, 19);
            this.ExpShow.TabIndex = 38;
            this.toolTip1.SetToolTip(this.ExpShow, "エクスプローラーで表示");
            this.ExpShow.UseVisualStyleBackColor = true;
            this.ExpShow.Visible = false;
            this.ExpShow.Click += new System.EventHandler(this.ExpShow_Click);
            // 
            // Folder
            // 
            this.Folder.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.Folder.FlatAppearance.BorderSize = 0;
            this.Folder.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.Folder.Font = new System.Drawing.Font("Meiryo UI", 6.75F);
            this.Folder.Image = ((System.Drawing.Image)(resources.GetObject("Folder.Image")));
            this.Folder.Location = new System.Drawing.Point(291, 39);
            this.Folder.Name = "Folder";
            this.Folder.Size = new System.Drawing.Size(24, 21);
            this.Folder.TabIndex = 18;
            this.toolTip1.SetToolTip(this.Folder, "フォルダの選択");
            this.Folder.UseVisualStyleBackColor = true;
            this.Folder.Click += new System.EventHandler(this.Folder_Click);
            // 
            // button6
            // 
            this.button6.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.button6.FlatAppearance.BorderSize = 0;
            this.button6.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button6.Font = new System.Drawing.Font("Meiryo UI", 6.75F);
            this.button6.Image = ((System.Drawing.Image)(resources.GetObject("button6.Image")));
            this.button6.Location = new System.Drawing.Point(270, 39);
            this.button6.Margin = new System.Windows.Forms.Padding(0);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(21, 21);
            this.button6.TabIndex = 15;
            this.button6.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.toolTip1.SetToolTip(this.button6, "リンクファイルからフォルダ情報を取得");
            this.button6.UseVisualStyleBackColor = true;
            this.button6.Click += new System.EventHandler(this.button6_Click);
            // 
            // show
            // 
            this.show.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.show.FlatAppearance.BorderSize = 0;
            this.show.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.show.Image = ((System.Drawing.Image)(resources.GetObject("show.Image")));
            this.show.Location = new System.Drawing.Point(276, 27);
            this.show.Name = "show";
            this.show.Size = new System.Drawing.Size(20, 19);
            this.show.TabIndex = 37;
            this.toolTip1.SetToolTip(this.show, "リンクへ移動");
            this.show.UseVisualStyleBackColor = true;
            this.show.Visible = false;
            this.show.Click += new System.EventHandler(this.button9_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.ErrorImage = null;
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.InitialImage = null;
            this.pictureBox1.Location = new System.Drawing.Point(72, 29);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(18, 18);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.pictureBox1.TabIndex = 33;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.Visible = false;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(69, 32);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(39, 14);
            this.label4.TabIndex = 35;
            this.label4.Text = "label4";
            // 
            // buttonApp4
            // 
            this.buttonApp4.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonApp4.Font = new System.Drawing.Font("Meiryo UI", 7F);
            this.buttonApp4.Location = new System.Drawing.Point(136, 4);
            this.buttonApp4.Margin = new System.Windows.Forms.Padding(0);
            this.buttonApp4.Name = "buttonApp4";
            this.buttonApp4.Size = new System.Drawing.Size(45, 23);
            this.buttonApp4.TabIndex = 31;
            this.buttonApp4.Text = "App4";
            this.buttonApp4.UseVisualStyleBackColor = true;
            this.buttonApp4.Click += new System.EventHandler(this.buttonApp4_Click);
            // 
            // buttonApp1
            // 
            this.buttonApp1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonApp1.Font = new System.Drawing.Font("Meiryo UI", 7F);
            this.buttonApp1.Location = new System.Drawing.Point(4, 4);
            this.buttonApp1.Margin = new System.Windows.Forms.Padding(0);
            this.buttonApp1.Name = "buttonApp1";
            this.buttonApp1.Size = new System.Drawing.Size(45, 23);
            this.buttonApp1.TabIndex = 28;
            this.buttonApp1.Text = "App1";
            this.buttonApp1.UseVisualStyleBackColor = true;
            this.buttonApp1.Click += new System.EventHandler(this.buttonApp1_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.textBox3);
            this.groupBox1.Controls.Add(this.Folder);
            this.groupBox1.Controls.Add(this.button2);
            this.groupBox1.Controls.Add(this.button5);
            this.groupBox1.Controls.Add(this.ReLink);
            this.groupBox1.Controls.Add(this.textBox2);
            this.groupBox1.Controls.Add(this.button6);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.pictureSame);
            this.groupBox1.Controls.Add(this.pictureBox2);
            this.groupBox1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.groupBox1.Font = new System.Drawing.Font("Meiryo UI", 8F);
            this.groupBox1.Location = new System.Drawing.Point(4, 102);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(317, 85);
            this.groupBox1.TabIndex = 34;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "再リンク先設定(コピー先はドラッグ＆ドロップも可)";
            // 
            // textBox3
            // 
            this.textBox3.AllowDrop = true;
            this.textBox3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox3.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.textBox3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBox3.Location = new System.Drawing.Point(44, 39);
            this.textBox3.Name = "textBox3";
            this.textBox3.ReadOnly = true;
            this.textBox3.Size = new System.Drawing.Size(226, 21);
            this.textBox3.TabIndex = 11;
            this.textBox3.TextChanged += new System.EventHandler(this.textBox2_TextChanged);
            this.textBox3.DragDrop += new System.Windows.Forms.DragEventHandler(this.textBox3_DragDrop);
            this.textBox3.DragEnter += new System.Windows.Forms.DragEventHandler(this.textBox3_DragEnter);
            // 
            // button2
            // 
            this.button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button2.Font = new System.Drawing.Font("Meiryo UI", 6.75F);
            this.button2.Location = new System.Drawing.Point(44, 62);
            this.button2.Margin = new System.Windows.Forms.Padding(0);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(108, 21);
            this.button2.TabIndex = 19;
            this.button2.Text = "リネームを行う";
            this.button2.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click_1);
            // 
            // button5
            // 
            this.button5.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button5.Font = new System.Drawing.Font("Meiryo UI", 6.75F);
            this.button5.Location = new System.Drawing.Point(154, 62);
            this.button5.Margin = new System.Windows.Forms.Padding(0);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(108, 21);
            this.button5.TabIndex = 14;
            this.button5.Text = "コピーのみ(再リンクしない)";
            this.button5.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // ReLink
            // 
            this.ReLink.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.ReLink.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.ReLink.Font = new System.Drawing.Font("Meiryo UI", 6.75F);
            this.ReLink.Location = new System.Drawing.Point(239, 17);
            this.ReLink.Margin = new System.Windows.Forms.Padding(0);
            this.ReLink.Name = "ReLink";
            this.ReLink.Size = new System.Drawing.Size(76, 21);
            this.ReLink.TabIndex = 12;
            this.ReLink.Text = "コピーして再リンク";
            this.ReLink.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.ReLink.UseVisualStyleBackColor = true;
            this.ReLink.Click += new System.EventHandler(this.button2_Click);
            // 
            // textBox2
            // 
            this.textBox2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox2.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.textBox2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBox2.Location = new System.Drawing.Point(44, 17);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(174, 21);
            this.textBox2.TabIndex = 8;
            this.textBox2.TextChanged += new System.EventHandler(this.textBox2_TextChanged);
            this.textBox2.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox2_KeyPress);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Meiryo UI", 7.5F);
            this.label2.Location = new System.Drawing.Point(1, 22);
            this.label2.Margin = new System.Windows.Forms.Padding(0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(45, 13);
            this.label2.TabIndex = 12;
            this.label2.Text = "ファイル名";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Meiryo UI", 7.5F);
            this.label3.Location = new System.Drawing.Point(1, 44);
            this.label3.Margin = new System.Windows.Forms.Padding(0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(39, 13);
            this.label3.TabIndex = 13;
            this.label3.Text = "コピー先";
            // 
            // pictureSame
            // 
            this.pictureSame.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.pictureSame.ErrorImage = null;
            this.pictureSame.Image = ((System.Drawing.Image)(resources.GetObject("pictureSame.Image")));
            this.pictureSame.InitialImage = null;
            this.pictureSame.Location = new System.Drawing.Point(219, 19);
            this.pictureSame.Name = "pictureSame";
            this.pictureSame.Size = new System.Drawing.Size(18, 18);
            this.pictureSame.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.pictureSame.TabIndex = 20;
            this.pictureSame.TabStop = false;
            this.pictureSame.Visible = false;
            // 
            // pictureBox2
            // 
            this.pictureBox2.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.pictureBox2.ErrorImage = null;
            this.pictureBox2.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox2.Image")));
            this.pictureBox2.InitialImage = null;
            this.pictureBox2.Location = new System.Drawing.Point(219, 19);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(18, 18);
            this.pictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.pictureBox2.TabIndex = 10;
            this.pictureBox2.TabStop = false;
            this.pictureBox2.Visible = false;
            // 
            // buttonApp2
            // 
            this.buttonApp2.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonApp2.Font = new System.Drawing.Font("Meiryo UI", 7F);
            this.buttonApp2.Location = new System.Drawing.Point(48, 4);
            this.buttonApp2.Margin = new System.Windows.Forms.Padding(0);
            this.buttonApp2.Name = "buttonApp2";
            this.buttonApp2.Size = new System.Drawing.Size(45, 23);
            this.buttonApp2.TabIndex = 29;
            this.buttonApp2.Text = "App2";
            this.buttonApp2.UseVisualStyleBackColor = true;
            this.buttonApp2.Click += new System.EventHandler(this.buttonApp2_Click);
            // 
            // buttonApp3
            // 
            this.buttonApp3.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonApp3.Font = new System.Drawing.Font("Meiryo UI", 7F);
            this.buttonApp3.Location = new System.Drawing.Point(92, 4);
            this.buttonApp3.Margin = new System.Windows.Forms.Padding(0);
            this.buttonApp3.Name = "buttonApp3";
            this.buttonApp3.Size = new System.Drawing.Size(45, 23);
            this.buttonApp3.TabIndex = 30;
            this.buttonApp3.Text = "App3";
            this.buttonApp3.UseVisualStyleBackColor = true;
            this.buttonApp3.Click += new System.EventHandler(this.buttonApp3_Click);
            // 
            // textBox1
            // 
            this.textBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox1.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.textBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBox1.Location = new System.Drawing.Point(4, 49);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBox1.Size = new System.Drawing.Size(315, 51);
            this.textBox1.TabIndex = 32;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(4, 32);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(70, 14);
            this.label5.TabIndex = 36;
            this.label5.Text = "ファイル情報：";
            // 
            // Form1
            // 
            this.AllowDrop = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Info;
            this.ClientSize = new System.Drawing.Size(326, 190);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.buttonShowILL);
            this.Controls.Add(this.Setting);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.buttonApp4);
            this.Controls.Add(this.buttonApp1);
            this.Controls.Add(this.ExpShow);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.show);
            this.Controls.Add(this.buttonApp2);
            this.Controls.Add(this.buttonApp3);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.label5);
            this.Font = new System.Drawing.Font("Meiryo UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimumSize = new System.Drawing.Size(290, 196);
            this.Name = "Form1";
            this.ShowInTaskbar = false;
            this.Text = "LillCop [β0.23]";
            this.TopMost = true;
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureSame)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Timer timer1;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button buttonShowILL;
        private System.Windows.Forms.Button Setting;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button buttonApp4;
        private System.Windows.Forms.Button buttonApp1;
        private System.Windows.Forms.Button ExpShow;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.PictureBox pictureSame;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.Button Folder;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.Button ReLink;
        private System.Windows.Forms.PictureBox pictureBox2;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button button6;
        private System.Windows.Forms.Button show;
        private System.Windows.Forms.Button buttonApp2;
        private System.Windows.Forms.Button buttonApp3;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label5;
    }
}

