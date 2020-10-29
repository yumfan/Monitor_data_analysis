using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
//using System.Linq;
using System.Text;
//using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;



namespace Battery_Soldering_Data_Collect
{
    public partial class Form1 : Form
    {
        //System.Collections.ArrayList alst;
        string selected_date;
        string path, path_n, path_p;
        string assigned_path;
        ArrayList file_list = new ArrayList();
        ArrayList file_list_neg = new ArrayList();
        ArrayList file_list_pos = new ArrayList();
        public Form1()
        {
            InitializeComponent();
            data_dispaly_grid();
            

        }

        private void data_dispaly_grid()
        {
            selected_date = Select_Date.Value.ToString("MM - dd - yyyy");    //得到当天的时间
            assigned_path = this.textBox1.Text;
            try
            {
                path = assigned_path + "synergy v5\\dati\\";
                path_n = assigned_path + "Rework_negative\\dati\\";
                path_p = assigned_path + "Rework_positive\\dati\\";
                file_list = GetFiles(path);
                file_list_neg = GetFiles(path_n);
                file_list_pos = GetFiles(path_p);
                //string[] sn_time;
                string[,] sn_time_result = new string[file_list.Count, 2];
                string[,] sn_time_result_neg = new string[file_list_neg.Count, 2];
                string[,] sn_time_result_pos = new string[file_list_neg.Count, 2];
                sn_time_result = get_sn_time(file_list);
                sn_time_result_neg = get_sn_time(file_list_neg);
                sn_time_result_pos = get_sn_time(file_list_pos);
                show_dataGrid(sn_time_result);
            }
            catch
            {
                return;
            }
        }


        private void CheckBox1_Clicked(object sender, EventArgs e)
        {
            this.checkBox1.Checked = true;
            this.checkBox2.Checked = false;
        }


        private void CheckBox2_Clicked(object sender, EventArgs e)
        {
            this.checkBox2.Checked = true;
            this.checkBox1.Checked = false;
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            int quantity = this.dataGridView1.Rows.Count;
            if (this.dataGridView1.Rows.Count <= 1)
            {
                MessageBox.Show("没有可供导出的数据");
                return;
            }
            else
            {
                if (checkBox1.Checked)
                {
                    save_excel(2);
                    //save_excel(0);

                    //System.Data.DataTable dt = (System.Data.DataTable)dataGridView1.DataSource;           //绑定datagrid的数据到dt
                    //dt.Rows.Clear();                                              //清除dt
                    //dataGridView1.DataSource = dt;                                //重新绑定清除后的dt到datagrid
                    string despath = this.textBox1.Text.ToString() + "\\result\\";
                    //MoveFiles(this.textBox1.Text + "synergy v5\\dati\\", despath);  //move 文件到\result目录下
                    //MoveFiles(this.textBox1.Text, despath);

                }
                else
                {
                    save_excel(1);
                }
            }
            MessageBox.Show("导出数据完成");
        }

        private void save_excel(int mode)
        {
            if (mode == 1)
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Excel files (*.xls)|*.xls";
                saveFileDialog.FilterIndex = 0;
                saveFileDialog.RestoreDirectory = true;
                saveFileDialog.CreatePrompt = true;
                saveFileDialog.Title = "导出Excel文件到";

                DateTime now = DateTime.Now;
                saveFileDialog.FileName = now.Year.ToString().PadLeft(2)
                + now.Month.ToString().PadLeft(2, '0')
                + now.Day.ToString().PadLeft(2, '0') + "-"
                + now.Hour.ToString().PadLeft(2, '0')
                + now.Minute.ToString().PadLeft(2, '0')
                + now.Second.ToString().PadLeft(2, '0');
                saveFileDialog.ShowDialog();

                Stream myStream = saveFileDialog.OpenFile();
                StreamWriter sw = new StreamWriter(myStream, System.Text.Encoding.GetEncoding("gb2312"));
                string str = "";
                try
                {
                    //写标题     
                    for (int i = 0; i < this.dataGridView1.ColumnCount; i++)
                    {
                        if (i > 0)
                        {
                            str += ",";
                        }
                        str += this.dataGridView1.Columns[i].HeaderText;
                    }
                    str = "\t" + str;
                    sw.WriteLine(str);
                    //写内容   
                    for (int j = 0; j < this.dataGridView1.Rows.Count; j++)
                    {
                        string tempStr = "";
                        for (int k = 0; k < this.dataGridView1.Columns.Count; k++)
                        {
                            if (k > 0)
                            {
                                tempStr += ",";
                            }
                            tempStr += this.dataGridView1.Rows[j].Cells[k].Value.ToString();
                        }
                        tempStr = (j + 1).ToString() + "," + tempStr;
                        sw.WriteLine(tempStr);
                    }
                    sw.Close();
                    myStream.Close();
                }
                catch (Exception)
                {
                    //MessageBox.Show(ex.ToString());
                }
                finally
                {
                    sw.Close();
                    myStream.Close();
                }
            }
            else if (mode == 0)
            {
                string[,] data = new string[this.dataGridView1.Rows.Count - 1, this.dataGridView1.Columns.Count];
                for (int i = 0; i < this.dataGridView1.Rows.Count - 1; i++)
                {
                    for (int j = 0; j < this.dataGridView1.Columns.Count; j++)
                    {
                        data[i, j] = this.dataGridView1.Rows[i].Cells[j].Value.ToString();
                    }
                }

                while (data[0, 0] != null)
                {
                    string[,] temp_same_date = new string[data.GetLength(0), data.GetLength(1)];
                    string[,] temp_diff_date = new string[data.GetLength(0), data.GetLength(1)];
                    int g = 0;
                    int h = 0;
                    for (int k = 0; k < data.GetLength(0); k++)
                    {
                        string date_temp = data[0, 1];

                        if (data[k, 1] != date_temp)
                        {
                            temp_diff_date[g, 0] = data[k, 0];
                            temp_diff_date[g, 1] = data[k, 1];
                            temp_diff_date[g, 2] = data[k, 2];
                            g++;
                        }
                        else
                        {
                            temp_same_date[h, 0] = data[k, 0];
                            temp_same_date[h, 1] = data[k, 1];
                            temp_same_date[h, 2] = data[k, 2];
                            h++;
                        }
                    }
                    data = new string[temp_diff_date.GetLength(0), temp_diff_date.GetLength(1)];
                    for (int k = 0; k < temp_diff_date.GetLength(0); k++)
                    {
                        data[k, 0] = temp_diff_date[k, 0];
                        data[k, 1] = temp_diff_date[k, 1];
                        data[k, 2] = temp_diff_date[k, 2];
                    }

                    string save_file_name = this.textBox1.Text + "\\" + "report\\" + temp_same_date[0, 1] + ".xls";
                    string save_file_path = this.textBox1.Text + "\\" + "report\\";
                    Directory.CreateDirectory(save_file_path);
                    if (!File.Exists(save_file_name))
                        File.Create(save_file_name).Close();

                    StreamWriter sw = new StreamWriter(save_file_name, false, Encoding.UTF8);
                    string str = "";
                    //写标题     
                    str = "," + "序列号" + "," + "日期" + "," + "时间";
                    sw.WriteLine(str);

                    //写内容
                    for (int j = 0; j < temp_same_date.GetLength(0); j++)
                    {
                        string tempStr = "";
                        for (int k = 0; k < temp_same_date.GetLength(1); k++)
                        {
                            if (k > 0)
                            {
                                tempStr += ",";
                            }
                            tempStr += temp_same_date[j, k];
                        }
                        if (tempStr != ",,")
                        {
                            tempStr = (j + 1).ToString() + "," + tempStr;
                            sw.WriteLine(tempStr);
                        }
                    }
                    sw.Flush();
                    sw.Close();

                }
            }
            else if (mode == 2)
            {
                System.Data.DataTable dt = (System.Data.DataTable)dataGridView1.DataSource;
                if (dt == null || dt.Rows.Count == 0) return;

                // 创建Excel文档，保存格式 Office2007 xlsx.
                // 需要引用：Microsoft.Office.Interop.Excel.dll 12.0版本的支持Office2007.
                while (dt.Rows[0][0] !="")
                {
                    string[,] temp_same_date = new string[dt.Rows.Count, dt.Columns.Count];          //定义2维数组长度
                    string[,] temp_diff_date = new string[dt.Rows.Count, dt.Columns.Count];         //定义2维数组长度
                    int g = 0;
                    int h = 0;
                    for (int k = 0; k < dt.Rows.Count; k++)
                    {
                        string date_temp = dt.Rows[0][1].ToString();
                        if (date_temp=="" )
                        {
                                return;
                            
                        }

                        if (dt.Rows[k][1].ToString() != date_temp)
                        {
                            temp_diff_date[g, 0] = dt.Rows[k][0].ToString();
                            temp_diff_date[g, 1] = dt.Rows[k][1].ToString();
                            temp_diff_date[g, 2] = dt.Rows[k][2].ToString();
                            g++;
                        }
                        else
                        {
                            temp_same_date[h, 0] = dt.Rows[k][0].ToString();
                            temp_same_date[h, 1] = dt.Rows[k][1].ToString();
                            temp_same_date[h, 2] = dt.Rows[k][2].ToString();
                            h++;
                        }
                    }
                    //temp_same_date = temp_same_date.Where(s => !string.IsNullOrEmpty(s)).ToArray();
                    dt.Clear();
                    for (int k = 0; k < temp_diff_date.GetLength(0); k++)    //将不同的数据重新写入表格中
                    {
                        DataRow dr = dt.NewRow();
                        for (int j = 0; j < temp_diff_date.GetLength(1); j++)
                            dr[j] = temp_diff_date[k, j];
                        dt.Rows.Add(dr);
                    }

                    /*
                    for (int k = 0; k < temp_diff_date.GetLength(0); k++)
                    {
                        dt.Rows[k][0] = temp_diff_date[k, 0];
                        dt.Rows[k][1] = temp_diff_date[k, 1];
                        dt.Rows[k][2] = temp_diff_date[k, 2];
                    }
                    */
                    string save_file_name = this.textBox1.Text + "\\" + "report\\" + temp_same_date[0, 1] + ".xlsx";
                    string save_file_path = this.textBox1.Text + "\\" + "report\\";
                    Directory.CreateDirectory(save_file_path);
                    /*
                    if (!File.Exists(save_file_name))
                        File.Create(save_file_name).Close();
                    else
                    {
                        System.IO.File.Delete(save_file_name);
                    }
                    */
                    if (File.Exists(save_file_name))
                    {                  
                       System.IO.File.Delete(save_file_name);      //如果原始excel文件存在，删除该文件
                    }
                    Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel._Workbook workBook = excel.Workbooks.Add();  //新建文件
                    Microsoft.Office.Interop.Excel._Worksheet workSheet = (Microsoft.Office.Interop.Excel._Worksheet)workBook.ActiveSheet; //新建sheet
                    object misValue = System.Reflection.Missing.Value;

                    //excel.DisplayAlerts = false; // 不显示告警

                    int rowIndex;
                    int colIndex;
                    int page;
                    Range range_1, range_2, range_3;
                    //取得标题并保存
                    String Current_date=DateTime.Now.ToString("yyyyMMdd");
                    String Number="记录单编码:  " + Current_date + "0001";
                    /*
                    workSheet.Cells[1, 1] = "电池序列号记录单";
                    Range range_1 = (Range)workSheet.get_Range("A1", "E1");
                    range_1.Merge(0);
                    range_1.HorizontalAlignment = XlVAlign.xlVAlignCenter;
                    workSheet.Cells[1, 6] = "记录单编码:  " + Current_date + "0001";
                    Range range_2 = (Range)workSheet.get_Range("F1", "I1");
                    range_2.Merge(0);
                    range_2.HorizontalAlignment = XlVAlign.xlVAlignCenter;
                    */
                    int same_date_length=0;
                    for (h=0;h< temp_same_date.GetLength(0);h++)
                    {
                        if(temp_same_date[h,0]!=null)
                        {
                            same_date_length++;
                        }
                    }
                    if (same_date_length % 100==0)
                    {
                        page = same_date_length / 100;
                    }
                    else
                    {
                        page = same_date_length / 100 + 1;
                    }
                    for (int k=0;k<page;k++)
                    {
                        rowIndex = 1;
                        colIndex = 0;
                        workSheet.Cells[54 + k * 58, 7] = "签名及日期:";
                        ///合并单元格
                        workSheet.Cells[1 + k * 58, 1] = "电池序列号记录单";
                        workSheet.Cells[1 + k * 58, 6] = Number;
                        range_1 = (Range)workSheet.get_Range("A" + Convert.ToString(1 + k * 58), "E" + Convert.ToString(1 + k * 58));
                        range_1.Merge(0);
                        range_1.HorizontalAlignment = XlVAlign.xlVAlignCenter;
                        range_2 = (Range)workSheet.get_Range("F" + Convert.ToString(1 + k * 58), "I" + Convert.ToString(1 + k * 58));
                        range_2.Merge(0);
                        range_2.HorizontalAlignment = XlVAlign.xlVAlignCenter;
                        for (int i = 0; i < dt.Columns.Count; i++)
                        {
                            colIndex++;
                            workSheet.Cells[2+k*58, colIndex+1] = dt.Columns[i].ColumnName;  //保存标题
                            workSheet.Cells[2+k*58, colIndex + 6] = dt.Columns[i].ColumnName;  //保存标题
                            workSheet.Cells[2+k*58, 1] = "序号";
                            workSheet.Cells[2+k*58, 6] = "序号";                                                       
                        }

                        //保存数据到execl
                        for (int i = k*100; i < (k+1) * 100; i++)
                        {
                            rowIndex++;
                            colIndex = 0;
                            /*
                            if (temp_same_date[i, 1] != null)              
                            {
                                workSheet.Cells[rowIndex+1, 1] = i + 1;        //保存数量编号到第一列
                            }
                        
                            */
                            if (i % 100 < 50 && i < same_date_length)
                            {
                                workSheet.Cells[rowIndex + 1 + k * 58, 1] = i + 1;
                            }

                            else
                                if (i < same_date_length)
                            {
                                workSheet.Cells[rowIndex + 1 + k * 58 - 50, 6] = i + 1;
                            
                            }
                            for (int j = 0; j < temp_same_date.GetLength(1); j++)
                            {

                                colIndex++;
                                if (i%100<50)
                                    {
                                        workSheet.Cells[rowIndex + 1 + k * 58, colIndex + 1] = temp_same_date[i, j];         //保存具体数据到excel                              
                                    }
                                else
                                    {
                                        workSheet.Cells[rowIndex + 1 + k * 58-50, colIndex + 6] = temp_same_date[i, j];
                                    }
                                    //workSheet.Cells.Borders.get_Item(XlBordersIndex.xlEdgeTop).LineStyle =Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                            }
                        }

                        //画边框，字体居中
                        range_3 = (Range)workSheet.get_Range("A" + Convert.ToString(1 + k * 58), "I" + Convert.ToString(52 + k * 58));
                        //横向居中
                        range_3.HorizontalAlignment = XlVAlign.xlVAlignCenter;
                        ///字体大小
                        range_3.Font.Size = 10;
                        ///字体
                        range_3.Font.Name = "黑体";
                        ///行高
                        //range_3.RowHeight = 24;
                        //自动调整列宽
                        range_3.EntireColumn.AutoFit();
                        //填充颜色
                        //range_3.Interior.ColorIndex = 20;
                        //设置单元格边框的粗细
                        range_3.Cells.Borders.LineStyle = 1;
                    }

                    

                    //文件保存
                    workSheet.SaveAs(save_file_name);      //保存文件
                    workBook.Close();                      //关闭引用
                    excel.Quit();                          //退出excel

                    PublicMethod.Kill(excel);//调用kill当前excel进程
                    releaseObject(workSheet);//释放COM对象
                    releaseObject(workBook);
                    releaseObject(excel);

                    GC.Collect();
                    //MessageBox.Show(string.Format("{0} 文档生成成功.", filePath));
                }
                return;
            }
        }

        private string SelectPath()
        {
            string path = string.Empty;
            System.Windows.Forms.FolderBrowserDialog fbd = new System.Windows.Forms.FolderBrowserDialog();
            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                path = fbd.SelectedPath;
            }
            return path;
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            string selected_path;
            selected_path = SelectPath();
            this.textBox1.Text = selected_path;
            string assigned_path = this.textBox1.Text;
            path = assigned_path + "\\synergy v5\\dati\\";
            path_n = assigned_path + "\\Rework_negative\\dati\\";
            path_p = assigned_path + "\\Rework_positive\\dati\\";
            file_list = GetFiles(path);
            file_list_neg = GetFiles(path_n);
            file_list_pos = GetFiles(path_p);
            //string[] sn_time;
            string[,] sn_time_result = new string[file_list.Count, 2];
            string[,] sn_time_result_neg = new string[file_list_neg.Count, 2];
            string[,] sn_time_result_pos = new string[file_list_neg.Count, 2];
            sn_time_result = get_sn_time(file_list);
            show_dataGrid(sn_time_result);
        }

        private void show_dataGrid(string[,] sn_time_result)
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            /*
            for (int i = 0; i < sn_time_result.GetLength(1); i++)
                dt.Columns.Add(i.ToString(), typeof(string));
            */
            dt.Columns.Add("序列号");
            dt.Columns.Add("焊接日期");               //加入表头
            dt.Columns.Add("焊接时间");
            dt.Columns.Add("Pass/Fail");
            for (int i = 0; i < sn_time_result.GetLength(0); i++)
            {
                DataRow dr = dt.NewRow();
                for (int j = 0; j < sn_time_result.GetLength(1); j++)
                    dr[j] = sn_time_result[i, j];
                dt.Rows.Add(dr);
            }
            this.dataGridView1.DataSource = dt;
        }



        private string[,] get_sn_time(ArrayList file_list)
        {
            try
            {
                // 创建一个 StreamReader 的实例来读取文件 
                // using 语句也能关闭 StreamReader
                string[,] array_sn_time = new string[file_list.Count, 4];
                string line, str_sn, str_date, str_time;
                int i = 0;
                foreach (string file in file_list)//循环文件

                {
                    using (StreamReader sr = new StreamReader(file))
                    {

                        line = sr.ReadLine();                    //读取一行-第一行
                        str_sn = line.Substring(0, 8);            //读取字符串起始0，长度8，sn
                        str_date = line.Substring(23, 10);       //读取字符串起始23，长度20，time
                        str_time = line.Substring(35, 8);       //读取字符串起始23，长度20，time
                        array_sn_time[i, 0] = str_sn;             //写入2维数组
                        array_sn_time[i, 1] = str_date;         //写入2维数组
                        array_sn_time[i, 2] = str_time;           //写入2维数组
                        //i++;                                     //移动2维数组index


                        // 从文件读取并显示行，直到文件的末尾 
                        /*
                        while ((line = sr.ReadLine()) != null)
                        {
                            Console.WriteLine(line);
                        }
                        */
                    }
                    array_sn_time[i, 3] = "Fail";
                    string[] ReadText = File.ReadAllLines(file, Encoding.Default);
                    foreach (string item in ReadText)
                    {
                        if (item.ToUpper().Contains("OK"))//搜索指定字符串ZZZ
                        {
                            array_sn_time[i, 3] = "OK";
                        }
                    }
                    i++;
                }
                return array_sn_time;
            }
            catch (Exception e)
            {
                // 向用户显示出错消息
                Console.WriteLine("The file could not be read:");
                Console.WriteLine(e.Message);
                return null;
            }
        }
        private string get_Date(string file_name)
        {
            try
            {
                // 创建一个 StreamReader 的实例来读取文件 
                // using 语句也能关闭 StreamReader
                string line, str_date;
                using (StreamReader sr = new StreamReader(file_name))
                    {

                        line = sr.ReadLine();                    //读取一行-第一行
                        
                        str_date = line.Substring(23, 10);       //读取字符串起始23，长度20，time
                        
                        // 从文件读取并显示行，直到文件的末尾 
                        /*
                        while ((line = sr.ReadLine()) != null)
                        {
                            Console.WriteLine(line);
                        }
                        */
                    }
                return str_date;
            }
            catch (Exception e)
            {
                // 向用户显示出错消息
                Console.WriteLine("The file could not be read:");
                Console.WriteLine(e.Message);
                return null;
            }
        }


        public ArrayList GetFiles(string dir)
        {
            try
            {
                string[] files = Directory.GetFiles(dir);//得到文件
                ArrayList file_list = new ArrayList();
                ArrayList file_time_list = new ArrayList();    //创建一个新的Arraylist
                foreach (string file in files)//循环文件
                {
                    string exname = file.Substring(file.LastIndexOf("_1.") + 1);//得到后缀名
                                                                                // if (".txt|.aspx".IndexOf(file.Substring(file.LastIndexOf(".") + 1)) > -1)//查找.txt .aspx结尾的文件
                                                                                //if (".DAT".IndexOf(file.Substring(file.LastIndexOf(".") + 1)) > -1)//如果后缀名为.dat文件
                    if (file.LastIndexOf("_1_1.DAT") > -1)//如果后缀名为.dat文件
                    {
                        FileInfo fi = new FileInfo(file);//建立FileInfo对象
                        file_list.Add(fi.FullName);//把.DAT文件全名加人到FileInfo对象
                        file_time_list.Add(fi.CreationTime);
                    }

                }
                return file_list;
            }
            catch
            {
                return null;
            }

        }

        private void Button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void MoveFiles(string srcFolder, string destFolder)
        {
            DirectoryInfo directoryInfo = new DirectoryInfo(srcFolder);
            FileInfo[] files = directoryInfo.GetFiles();
            String Today=DateTime.Now.ToString("MM-dd-yyyy");
            string Date;

            foreach (FileInfo file in files) // Directory.GetFiles(srcFolder)
            {
                if (file.Extension == ".DAT")
                {
                    Date = get_Date(file.FullName);
                    if (File.Exists(Path.Combine(destFolder, file.Name)))
                    {
                        System.IO.File.Delete(Path.Combine(destFolder, file.Name));      //如果原始log文件存在，删除该文件
                    }
                    if (Today != Date)
                    {
                        file.MoveTo(Path.Combine(destFolder, file.Name));
                        
                    }
                    }
                // will move all files without if stmt 
                //file.MoveTo(Path.Combine(destFolder, file.Name));
            }
        }
        private static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch
            {
                obj = null;
            }
        }

        public class PublicMethod
        {
            [DllImport("User32.dll", CharSet = CharSet.Auto)]
            public static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);

            public static void Kill(Microsoft.Office.Interop.Excel.Application excel)
            {
                try
                {
                    IntPtr ptr = new IntPtr(excel.Hwnd);
                    int processID = 0;

                    GetWindowThreadProcessId(ptr, out processID);
                    System.Diagnostics.Process process = System.Diagnostics.Process.GetProcessById(processID);
                    process.Kill();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(" ==== " + ex.Message);
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void DateTimePicker_1_ValueChanged(object sender, EventArgs e)
        {
            selected_date = Select_Date.Value.ToString("MM - dd - yyyy");
            return;
        }
    }
}
