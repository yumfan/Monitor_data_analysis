namespace Battery_Soldering_Data_Collect
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
            this.label1 = new System.Windows.Forms.Label();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.checkBox2 = new System.Windows.Forms.CheckBox();
            this.Generate_Report = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.Select_Path = new System.Windows.Forms.Button();
            this.Exit = new System.Windows.Forms.Button();
            this.Select_Date = new System.Windows.Forms.DateTimePicker();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(236, 74);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowHeadersWidth = 51;
            this.dataGridView1.RowTemplate.Height = 23;
            this.dataGridView1.Size = new System.Drawing.Size(479, 320);
            this.dataGridView1.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(236, 55);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(136, 16);
            this.label1.TabIndex = 1;
            this.label1.Text = "序列号-焊接时间";
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Enabled = false;
            this.checkBox1.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.checkBox1.Location = new System.Drawing.Point(37, 185);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(167, 16);
            this.checkBox1.TabIndex = 2;
            this.checkBox1.Text = "按焊接时间生成每日报告";
            this.checkBox1.UseVisualStyleBackColor = true;
            this.checkBox1.Visible = false;
            this.checkBox1.Click += new System.EventHandler(this.CheckBox1_Clicked);
            // 
            // checkBox2
            // 
            this.checkBox2.AutoSize = true;
            this.checkBox2.Checked = true;
            this.checkBox2.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox2.Enabled = false;
            this.checkBox2.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.checkBox2.Location = new System.Drawing.Point(37, 218);
            this.checkBox2.Name = "checkBox2";
            this.checkBox2.Size = new System.Drawing.Size(154, 16);
            this.checkBox2.TabIndex = 3;
            this.checkBox2.Text = "将选中的数据生成报告";
            this.checkBox2.UseVisualStyleBackColor = true;
            this.checkBox2.Click += new System.EventHandler(this.CheckBox2_Clicked);
            // 
            // Generate_Report
            // 
            this.Generate_Report.Location = new System.Drawing.Point(37, 255);
            this.Generate_Report.Name = "Generate_Report";
            this.Generate_Report.Size = new System.Drawing.Size(168, 23);
            this.Generate_Report.TabIndex = 4;
            this.Generate_Report.Text = "生成测试报告";
            this.Generate_Report.UseVisualStyleBackColor = true;
            this.Generate_Report.Click += new System.EventHandler(this.Button1_Click);
            // 
            // textBox1
            // 
            this.textBox1.Enabled = false;
            this.textBox1.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.textBox1.Location = new System.Drawing.Point(239, 412);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(305, 26);
            this.textBox1.TabIndex = 5;
            this.textBox1.Text = "C:\\Users\\fym12\\Documents\\创领心脏起搏器\\Production Testbed\\battery soldering_1\\GETSOLDER" +
    "_2\\";
            // 
            // Select_Path
            // 
            this.Select_Path.Location = new System.Drawing.Point(564, 412);
            this.Select_Path.Name = "Select_Path";
            this.Select_Path.Size = new System.Drawing.Size(126, 23);
            this.Select_Path.TabIndex = 6;
            this.Select_Path.Text = "选择log存储路径";
            this.Select_Path.UseVisualStyleBackColor = true;
            this.Select_Path.Click += new System.EventHandler(this.Button2_Click);
            // 
            // Exit
            // 
            this.Exit.Location = new System.Drawing.Point(37, 318);
            this.Exit.Margin = new System.Windows.Forms.Padding(2);
            this.Exit.Name = "Exit";
            this.Exit.Size = new System.Drawing.Size(168, 23);
            this.Exit.TabIndex = 7;
            this.Exit.Text = "退出";
            this.Exit.UseVisualStyleBackColor = true;
            this.Exit.Click += new System.EventHandler(this.Button3_Click);
            // 
            // Select_Date
            // 
            this.Select_Date.Location = new System.Drawing.Point(37, 142);
            this.Select_Date.Margin = new System.Windows.Forms.Padding(2);
            this.Select_Date.Name = "Select_Date";
            this.Select_Date.Size = new System.Drawing.Size(151, 21);
            this.Select_Date.TabIndex = 8;
            this.Select_Date.ValueChanged += new System.EventHandler(this.DateTimePicker_1_ValueChanged);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.Select_Date);
            this.Controls.Add(this.Exit);
            this.Controls.Add(this.Select_Path);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.Generate_Report);
            this.Controls.Add(this.checkBox2);
            this.Controls.Add(this.checkBox1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dataGridView1);
            this.Name = "Form1";
            this.Text = "Battery Soldering Data_V1.0";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private void CheckBox1_Click(object sender, System.EventArgs e)
        {
            throw new System.NotImplementedException();
        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.CheckBox checkBox2;
        private System.Windows.Forms.Button Generate_Report;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button Select_Path;
        private System.Windows.Forms.Button Exit;
        private System.Windows.Forms.DateTimePicker Select_Date;
    }
}

