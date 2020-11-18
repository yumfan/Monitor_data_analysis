namespace 起搏器_程控仪_监控数据_收集分析
{
    partial class Form1
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
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.folder = new System.Windows.Forms.TextBox();
            this.Select_Folder = new System.Windows.Forms.Button();
            this.retrieve_data = new System.Windows.Forms.Button();
            this.data_analyse = new System.Windows.Forms.Button();
            this.Exit = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCells;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(269, 28);
            this.dataGridView1.Margin = new System.Windows.Forms.Padding(2);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 27;
            this.dataGridView1.Size = new System.Drawing.Size(895, 511);
            this.dataGridView1.TabIndex = 0;
            // 
            // folder
            // 
            this.folder.Location = new System.Drawing.Point(11, 196);
            this.folder.Margin = new System.Windows.Forms.Padding(2);
            this.folder.Name = "folder";
            this.folder.Size = new System.Drawing.Size(244, 21);
            this.folder.TabIndex = 1;
            // 
            // Select_Folder
            // 
            this.Select_Folder.Location = new System.Drawing.Point(51, 241);
            this.Select_Folder.Margin = new System.Windows.Forms.Padding(2);
            this.Select_Folder.Name = "Select_Folder";
            this.Select_Folder.Size = new System.Drawing.Size(154, 29);
            this.Select_Folder.TabIndex = 2;
            this.Select_Folder.Text = "选择起搏器数据文件目录";
            this.Select_Folder.UseVisualStyleBackColor = true;
            this.Select_Folder.Click += new System.EventHandler(this.Select_Folder_Click);
            // 
            // retrieve_data
            // 
            this.retrieve_data.AutoSize = true;
            this.retrieve_data.Location = new System.Drawing.Point(51, 293);
            this.retrieve_data.Margin = new System.Windows.Forms.Padding(2);
            this.retrieve_data.Name = "retrieve_data";
            this.retrieve_data.Size = new System.Drawing.Size(154, 29);
            this.retrieve_data.TabIndex = 3;
            this.retrieve_data.Text = "检索数据_保存数据库";
            this.retrieve_data.UseVisualStyleBackColor = true;
            // 
            // data_analyse
            // 
            this.data_analyse.AutoSize = true;
            this.data_analyse.Location = new System.Drawing.Point(51, 351);
            this.data_analyse.Margin = new System.Windows.Forms.Padding(2);
            this.data_analyse.Name = "data_analyse";
            this.data_analyse.Size = new System.Drawing.Size(154, 30);
            this.data_analyse.TabIndex = 4;
            this.data_analyse.Text = "数据分析_筛选(数据库）";
            this.data_analyse.UseVisualStyleBackColor = true;
            // 
            // Exit
            // 
            this.Exit.AutoSize = true;
            this.Exit.Location = new System.Drawing.Point(96, 410);
            this.Exit.Margin = new System.Windows.Forms.Padding(2);
            this.Exit.Name = "Exit";
            this.Exit.Size = new System.Drawing.Size(56, 26);
            this.Exit.TabIndex = 5;
            this.Exit.Text = "退出";
            this.Exit.UseVisualStyleBackColor = true;
            this.Exit.Click += new System.EventHandler(this.Exit_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1175, 602);
            this.Controls.Add(this.Exit);
            this.Controls.Add(this.data_analyse);
            this.Controls.Add(this.retrieve_data);
            this.Controls.Add(this.Select_Folder);
            this.Controls.Add(this.folder);
            this.Controls.Add(this.dataGridView1);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Main";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.TextBox folder;
        private System.Windows.Forms.Button Select_Folder;
        private System.Windows.Forms.Button retrieve_data;
        private System.Windows.Forms.Button data_analyse;
        private System.Windows.Forms.Button Exit;
    }
}

