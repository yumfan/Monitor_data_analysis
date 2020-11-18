using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace 起搏器_程控仪_监控数据_收集分析
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            this.WindowState = FormWindowState.Normal;
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

        private void Select_Folder_Click(object sender, EventArgs e)
        {
            string selected_path;
            selected_path = SelectPath();
            this.folder.Text = selected_path;
            this.data_summary();
        }

        private void Exit_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }

        private void data_summary()
        {
            //selected_date = Select_Date.Value.ToString("MM-dd-yyyy");    //得到当天的时间
            string assigned_path = this.folder.Text;
            ArrayList overview_file_list = new ArrayList();
            ArrayList patient_file_list = new ArrayList();
            ArrayList cloumns_head = new ArrayList();
            int overview_file_count = Get_overview_file_count(assigned_path,"\\overview");
            string[,] full_data = new string[overview_file_count,40];
            try
            {
                string[] folders = Directory.GetDirectories(assigned_path);//得到每次保存的子目录
                int data_rows = 0;
                foreach (string sub_folder in folders)
                {
                    data_rows++;
                    overview_file_list = GetFiles(sub_folder, "\\overview\\");                 //得到一个子目录里的overview文件名
                    patient_file_list = GetFiles(sub_folder, "\\Patient\\");                   //得到一个子目录里的patient文件名
                    ArrayList data_item = excel_item_read(overview_file_list[0].ToString());   //得到一个report里的overiew标题
                    ArrayList patient_item = excel_item_read(patient_file_list[0].ToString()); //得到一个report里的patient标题
                    cloumns_head = head_combinate(data_item, patient_item);  //合并显示的标题名称                    
                    ArrayList patient_data = excel_data_read(patient_file_list[0].ToString());  //patient所有文件的内容是一样的，只取第一个文件
                    foreach (string file_name in overview_file_list)//循环顺序编号目录 11601/11602....
                    {
                        ArrayList overview_data = excel_data_read(file_name);
                        ArrayList data = data_combinate(overview_data,patient_data);
                        for(int i=0;i<data_rows;i++)
                        {
                            for(int j=0;j<data.Count;j++)
                            {
                                full_data[i, j] = data[j].ToString();
                            }
                        }
                    }
                    
                }
                show_dataGrid(cloumns_head,full_data);
                //string[] sn_time;
                //string[,] sn_time_result_1 = new string[file_list_1.Count, 2];
                //sn_time_result_1 = get_sn_time(file_list_1, selected_date, "normal", "正极");
                //show_dataGrid(sn_time_result_1, sn_time_result_2, sn_time_result_neg_1, sn_time_result_neg_2, sn_time_result_pos_1, sn_time_result_pos_2, sn_time_result_1_2, sn_time_result_2_2);
            }
            catch
            {
                return;
            }
        }
        public ArrayList head_combinate(ArrayList data,ArrayList patient)
        {
            ArrayList head = new ArrayList();
            for (int i=0;i<patient.Count;i++)
            {
                if (i == 0 || i == 1 || i == 23 || i == 24 || i == 25)
                {
                    head.Add(patient[i]);
                }
            }
            head.Add("植入日期");
            head.Add("数据采集日期");
            for (int i = 0; i < data.Count; i++)
            {
                if (data[i].ToString().LastIndexOf("日")==-1 && data[i].ToString().LastIndexOf("月") == -1 && data[i].ToString().LastIndexOf("年") == -1)
                    {
                       //包含年月日的不要
                       head.Add(data[i]);
                    }
            }
            return head;
        }
        public ArrayList data_combinate(ArrayList overview_data, ArrayList patient_data)
        {
            ArrayList final_data = new ArrayList();
            string temp_data;
            for (int i = 0; i < patient_data.Count; i++)
            {
                if (i == 0 || i == 1 || i == 23 || i == 24 || i == 25)
                {
                    final_data.Add(patient_data[i]);
                }
            }
            final_data.Add(patient_data[28]+"."+patient_data[27]+"."+patient_data[26]);
            final_data.Add(overview_data[50]+"."+ overview_data[49]+"."+ overview_data[48]);
            for (int i = 0; i < overview_data.Count; i++)
            {   
                if(overview_data[i]==null)
                {
                    temp_data = "NA";
                }
                else
                {
                    temp_data = overview_data[i].ToString();
                }
                if (temp_data!=overview_data[50].ToString()  && temp_data != overview_data[49].ToString() && temp_data != overview_data[48].ToString())
                {
                    //包含年月日的不要
                   final_data.Add(temp_data);
                }
            }
            return final_data;
        }
        public ArrayList GetFiles(string folder,string file_type)
        {
            try
            {
                string[] files;
                //string[] folders = Directory.GetDirectories(dir);//得到文件
                ArrayList file_list = new ArrayList();
                //foreach (string folder in folders)//循环顺序编号目录 11601/11602....
                //{
                    files = Directory.GetFiles(folder + file_type);
                    //string exname = file.Substring(file.LastIndexOf("_1.") + 1);//得到后缀名
                    // if (".txt|.aspx".IndexOf(file.Substring(file.LastIndexOf(".") + 1)) > -1)//查找.txt .aspx结尾的文件
                    //if (".DAT".IndexOf(file.Substring(file.LastIndexOf(".") + 1)) > -1)//如果后缀名为.dat文件

                    foreach (string file in files)//循环顺序编号目录 11601/11602....
                    {
                        if (file.LastIndexOf(".xlsx") > -1&& file.LastIndexOf("$")==-1) //如果后缀名为.xksx文件，包含$字符的临时文件必须排除
                        {
                            FileInfo fi = new FileInfo(file);//建立FileInfo对象
                            //string date = fi.CreationTime.ToString("MM-dd-yyyy");
                            file_list.Add(fi.FullName);//把.DAT文件全名加人到FileInfo对象
                            //file_time_list.Add(fi.CreationTime);


                        }
                    }
                    
                //}
                return file_list;
            }
            catch
            {
                return null;
            }

        }

        public int Get_overview_file_count(string dir, string file_type)
        {
            try
            {
                string[] files;
                string[] folders = Directory.GetDirectories(dir);//得到文件
                ArrayList file_list = new ArrayList();
                foreach (string folder in folders)//循环顺序编号目录 11601/11602....
                {
                files = Directory.GetFiles(folder + file_type);
                //string exname = file.Substring(file.LastIndexOf("_1.") + 1);//得到后缀名
                // if (".txt|.aspx".IndexOf(file.Substring(file.LastIndexOf(".") + 1)) > -1)//查找.txt .aspx结尾的文件
                //if (".DAT".IndexOf(file.Substring(file.LastIndexOf(".") + 1)) > -1)//如果后缀名为.dat文件

                    foreach (string file in files)//循环顺序编号目录 11601/11602....
                    {
                        if (file.LastIndexOf(".xlsx") > -1 && file.LastIndexOf("$") == -1) //如果后缀名为.xksx文件，包含$字符的临时文件必须排除
                        {
                                FileInfo fi = new FileInfo(file);//建立FileInfo对象
                                                         //string date = fi.CreationTime.ToString("MM-dd-yyyy");
                                file_list.Add(fi.FullName);//把.DAT文件全名加人到FileInfo对象
                                                   //file_time_list.Add(fi.CreationTime);
                        }
                    }

                }
                return file_list.Count;
            }
            catch
            {
                return 0;
            }

        }

        private void show_dataGrid(ArrayList head_name,string[,] full_data )
        {
            System.Data.DataTable dt = new System.Data.DataTable();        //实例dt维datatable
            /*
            for (int i = 0; i < sn_time_result.GetLength(1); i++)
                dt.Columns.Add(i.ToString(), typeof(string));
            */
            dt.Columns.Add("编号");
            for (int i = 0; i < head_name.Count; i++)
            {
                dt.Columns.Add(head_name[i].ToString());//加入表头
            }
            int rows = 1;
            for (int i=0;i<full_data.GetLength(0); i++)
            {
                DataRow dr = dt.NewRow();
                dr[0] = rows++;
                for(int j=0;j<full_data.GetLength(1);j++)
                {
                    dr[j+1] = full_data[i, j];
                }
                dt.Rows.Add(dr);
            }

           /*
            
            int count = 1;
            for (int i = 0; i < result_1.GetLength(0); i++)
            {
                DataRow dr = dt.NewRow();                               //实例dt的行
                if (result_1[i, 0] != null)
                {
                    for (int h = 0; h < result_2.GetLength(0); h++)
                    {
                        if (result_1[i, 0] == result_2[h, 0] && result_1[i, 1] == result_2[h, 1])
                        {
                            if (result_1[i, 4] != "Pass-正极" && result_2[h, 4] != "Pass-负极")
                            {
                                result_1[i, 4] = "Fail_正极&负极";
                            }
                            else
                            {
                                if (result_1[i, 4] == "Pass-正极" && result_2[h, 4] == "Pass-负极")
                                {
                                    result_1[i, 4] = "Pass";
                                }
                                if (result_2[h, 4] != "Pass-负极")
                                {
                                    result_1[i, 4] = "Fail_负极";
                                }
                            }

                        }
                    }
                    for (int j = 0; j < result_1.GetLength(1) + 1; j++)
                        if (j == 0)
                        {
                            dr[j] = count++;
                        }
                        else
                        {
                            dr[j] = result_1[i, j - 1];
                        }
                    dt.Rows.Add(dr);
                }
            }
            for (int i = 0; i < result_1_2.GetLength(0); i++)
            {
                DataRow dr = dt.NewRow();                               //实例dt的行
                if (result_1_2[i, 0] != null)
                {
                    for (int j = 0; j < result_1_2.GetLength(1) + 1; j++)
                        if (j == 0)
                        {
                            dr[j] = count++;
                        }
                        else
                        {
                            dr[j] = result_1_2[i, j - 1];
                        }
                    dt.Rows.Add(dr);
                }
            }
            for (int i = 0; i < result_2_2.GetLength(0); i++)
            {
                DataRow dr = dt.NewRow();                               //实例dt的行
                if (result_2_2[i, 0] != null)
                {
                    for (int j = 0; j < result_2_2.GetLength(1) + 1; j++)
                        if (j == 0)
                        {
                            dr[j] = count++;
                        }
                        else
                        {
                            dr[j] = result_2_2[i, j - 1];
                        }
                    dt.Rows.Add(dr);
                }
            }
            for (int i = 0; i < result_neg_1.GetLength(0); i++)
            {
                DataRow dr = dt.NewRow();                               //实例dt的行
                if (result_neg_1[i, 0] != null)
                {
                    for (int j = 0; j < result_neg_1.GetLength(1) + 1; j++)
                        if (j == 0)
                        {
                            dr[j] = count++;
                        }
                        else
                        {
                            dr[j] = result_neg_1[i, j - 1];
                        }
                    dt.Rows.Add(dr);
                }
            }
            for (int i = 0; i < result_neg_2.GetLength(0); i++)
            {
                DataRow dr = dt.NewRow();                               //实例dt的行
                if (result_neg_2[i, 0] != null)
                {
                    for (int j = 0; j < result_neg_2.GetLength(1) + 1; j++)
                        if (j == 0)
                        {
                            dr[j] = count++;
                        }
                        else
                        {
                            dr[j] = result_neg_2[i, j - 1];
                        }
                    dt.Rows.Add(dr);
                }
            }
            for (int i = 0; i < result_pos_1.GetLength(0); i++)
            {
                DataRow dr = dt.NewRow();                               //实例dt的行
                if (result_pos_1[i, 0] != null)
                {
                    for (int j = 0; j < result_pos_1.GetLength(1) + 1; j++)
                        if (j == 0)
                        {
                            dr[j] = count++;
                        }
                        else
                        {
                            dr[j] = result_pos_1[i, j - 1];
                        }
                    dt.Rows.Add(dr);
                }
            }
            for (int i = 0; i < result_pos_2.GetLength(0); i++)
            {
                DataRow dr = dt.NewRow();                               //实例dt的行
                if (result_pos_2[i, 0] != null)
                {
                    for (int j = 0; j < result_pos_2.GetLength(1) + 1; j++)
                        if (j == 0)
                        {
                            dr[j] = count++;
                        }
                        else
                        {
                            dr[j] = result_pos_2[i, j - 1];
                        }
                    dt.Rows.Add(dr);
                }
            }
            */
            this.dataGridView1.DataSource = dt;
        }

        private string get_data(string file_name)
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

        private void save_excel()
        {
                        
                        
                System.Data.DataTable dt = (System.Data.DataTable)dataGridView1.DataSource;
                if (dt == null || dt.Rows.Count == 0) return;  //判断表格里数据是否为空

                // 创建Excel文档，保存格式 Office2007 xlsx.
                // 需要引用：Microsoft.Office.Interop.Excel.dll 12.0版本的支持Office2007.


                string save_file_path = this.folder.Text + "\\" + "report\\";
                Directory.CreateDirectory(save_file_path);
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
                String Current_date = DateTime.Now.ToString("yyyyMMdd");
                String Number = "记录单编码:  " + Current_date + "0001";
                string save_file_name = this.folder.Text + "report\\" + Current_date + ".xlsx";
                if (dt.Rows.Count % 100 == 0)
                {
                    page = dt.Rows.Count / 100;
                }
                else
                {
                    page = dt.Rows.Count / 100 + 1;
                }
                for (int k = 0; k < page; k++)
                {
                    rowIndex = 1;
                    colIndex = 0;
                    workSheet.Cells[54 + k * 58, 7] = "签名及日期:";
                    ///合并单元格
                    workSheet.Cells[1 + k * 58, 1] = "电池序列号记录单";
                    workSheet.Cells[1 + k * 58, 8] = Number;
                    range_1 = (Range)workSheet.get_Range("A" + Convert.ToString(1 + k * 58), "G" + Convert.ToString(1 + k * 58));
                    range_1.Merge(0);
                    range_1.HorizontalAlignment = XlVAlign.xlVAlignCenter;
                    range_2 = (Range)workSheet.get_Range("H" + Convert.ToString(1 + k * 58), "M" + Convert.ToString(1 + k * 58));
                    range_2.Merge(0);
                    range_2.HorizontalAlignment = XlVAlign.xlVAlignCenter;
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        colIndex++;
                        workSheet.Cells[2 + k * 58, colIndex + 0] = dt.Columns[i].ColumnName;  //保存标题
                        workSheet.Cells[2 + k * 58, colIndex + 7] = dt.Columns[i].ColumnName;  //保存标题
                                                                                               //workSheet.Cells[2 + k * 58, 1] = "序号";
                                                                                               //workSheet.Cells[2 + k * 58, 8] = "序号";
                    }
                    //画边框，字体居中
                    range_3 = (Range)workSheet.get_Range("A" + Convert.ToString(1 + k * 58), "M" + Convert.ToString(52 + k * 58));
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
                //保存数据到execl
                rowIndex = 1;
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    rowIndex++;
                    colIndex = 0;
                    if (i % 100 < 50)
                    {
                        workSheet.Cells[rowIndex % 102 + 1 + i / 100 * 60, 1] = i + 1;
                    }
                    else
                    {
                        workSheet.Cells[rowIndex % 102 + 1 + i / 100 * 60 - 50, 8] = i + 1;
                    }
                    if (i < dt.Rows.Count)
                    {
                        for (int j = 1; j < dt.Columns.Count; j++)
                        {
                            colIndex++;
                            if (i % 100 < 50)
                            {
                                workSheet.Cells[rowIndex % 102 + 1 + i / 100 * 60, colIndex + 1] = dt.Rows[i][j].ToString();         //保存具体数据到excel                              
                            }
                            else
                            {
                                workSheet.Cells[rowIndex % 102 + 1 + i / 100 * 60 - 50, colIndex + 8] = dt.Rows[i][j].ToString();
                                //workSheet.Cells.Borders.get_Item(XlBordersIndex.xlEdgeTop).LineStyle =Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                            }
                        }
                    }
                }
                //文件保存
                try
                {
                    if (File.Exists(save_file_name))
                    {
                        System.IO.File.Delete(save_file_name);      //如果原始excel文件存在，删除该文件
                    }
                    range_3 = (Range)workSheet.get_Range("A1", "M500");
                    //自动调整列宽
                    range_3.EntireColumn.AutoFit();
                    workSheet.SaveAs(save_file_name);      //保存文件
                    workBook.Close();                      //关闭引用
                    excel.Quit();                          //退出excel
                    PublicMethod.Kill(excel);//调用kill当前excel进程
                    releaseObject(workSheet);//释放COM对象
                    releaseObject(workBook);
                    releaseObject(excel);
                    GC.Collect();
                    MessageBox.Show("导出数据完成");
                }
                catch
                {
                    MessageBox.Show("无法保存文件,请先关闭Excel程序");
                }

                //MessageBox.Show(string.Format("{0} 文档生成成功."));
            
            return;
        }

        private ArrayList excel_item_read(string file_name)
        {

            ArrayList data_item=new ArrayList();
            object missing = System.Reflection.Missing.Value;

            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application(); //lauch excel application

            if (file_name == null)

            {

                MessageBox.Show("选定目录中没有测试数据Excel文件");
                return null;

            }

            else

            {

                excel.Visible = false; excel.UserControl = true;

                // 以只读的形式打开EXCEL文件

                Microsoft.Office.Interop.Excel.Workbook wb = excel.Application.Workbooks.Open(file_name, missing, true, missing, missing, missing, missing, missing, missing, true, missing, missing, missing, missing, missing);

                //取得第一个工作薄

                Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets.get_Item(1);

                //取得总记录行数    (包括标题列)

                int rowsint = ws.UsedRange.Cells.Rows.Count; //得到行数
                for (int i=2;i<rowsint+1;i++)
                {
                    data_item.Add(ws.Cells[i, 1].Value);
                }
                
                excel.Quit();                          //退出excel
                PublicMethod.Kill(excel);//调用kill当前excel进程
                releaseObject(ws);//释放COM对象
                releaseObject(wb);
                releaseObject(excel);
                GC.Collect();
                
            }            
            return data_item;
        }

        private ArrayList excel_data_read(string file_name)
        {

            ArrayList data=new ArrayList();
            object missing = System.Reflection.Missing.Value;

            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application(); //lauch excel application

            if (file_name == null)

            {

                MessageBox.Show("选定目录中没有测试数据Excel文件");
                return null;

            }

            else

            {

                excel.Visible = false; excel.UserControl = true;

                // 以只读的形式打开EXCEL文件

                Microsoft.Office.Interop.Excel.Workbook wb = excel.Application.Workbooks.Open(file_name, missing, true, missing, missing, missing, missing, missing, missing, true, missing, missing, missing, missing, missing);

                //取得第一个工作薄

                Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets.get_Item(1);

                //取得总记录行数    (包括标题列)

                int rowsint = ws.UsedRange.Cells.Rows.Count; //得到行数
                for (int i = 2; i < rowsint + 1; i++)
                    {
                        data.Add(ws.Cells[i, 3].Value);
                    }

                
                excel.Quit();                          //退出excel
                PublicMethod.Kill(excel);//调用kill当前excel进程
                releaseObject(ws);//释放COM对象
                releaseObject(wb);
                releaseObject(excel);
                GC.Collect();
                
            }
            return data;
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


    }
}
