using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Runtime.InteropServices.ComTypes;
using System.Drawing.Imaging;
using System.Linq.Expressions;
using ClosedXML.Excel;
using System.Reflection;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Timers;
using System.Diagnostics;
using DevExpress.XtraReports.Parameters;
using DocumentFormat.OpenXml.Bibliography;
using System.Runtime.InteropServices;
using System.Data.Odbc;
using System.Web;
using DevExpress.Utils.Filtering.Internal;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;
using System.Data;

namespace demo05
{
    public partial class UserControl1 : UserControl
    {
        //Connection string
        private string strCon;
        private SqlConnection sqlCon = new SqlConnection();
        public UserControl1()
        {
            InitializeComponent();
        }
        private void UserControl1_Load(object sender, EventArgs e)
        {

            btn_time1.Show();
            btn_time2.Show();
            //// Lấy ngày hiện tại
            DateTime currentDate = DateTime.Now;
            //// Lấy ngày tiếp theo
            DateTime nextDate = currentDate.AddDays(1);

            //// Đặt giá trị cho DateTime Picker
            btn_time2.Value = nextDate;
            btn_time1.Value = currentDate;

            //bảng setting bị ẩn

            groupBox2.Visible = false;
            group_load.Visible = false;
            group_save.Visible = false;

            //load du lieu
            textBox2.Text = "50";  //hiện trên textbox số record 
            string filePath = Path.Combine(@"C:\Report", "test.txt"); // Đường dẫn từ textbox
            tbox_load.Text= filePath;
            try
            {
                if (File.Exists(filePath))
                {
                    using (StreamReader reader = new StreamReader(filePath))
                    {
                        while (!reader.EndOfStream)
                        {
                            string line = reader.ReadLine();
                            string[] parts = line.Split(':'); // Tách chuỗi bằng dấu :

                            if (parts.Length >= 2)
                            {
                                string field = parts[0].Trim(); // Lấy phần trước dấu :
                                string value = string.Join(":", parts.Skip(1)).Trim(); // Lấy phần sau dấu :

                                switch (field)
                                {
                                    case "Server":
                                        tbox_server.Text = value;
                                        break;
                                    case "Database":
                                        tbox_data.Text = value;
                                        break;
                                    case "User":
                                        tbox_user.Text = value;
                                        break;
                                    case "Password":
                                        tbox_pass.Text = value;
                                        break;
                                    case "Source":
                                        tbox_source.Text = value;
                                        break;
                                    case "Export":
                                        tbox_export.Text = value;
                                        break;
                                    case "Table Name":
                                        combox_table.Text = value;
                                        break;
                                    case "Authentication":
                                        combox_authentication.Text = value;
                                        break;
                                    default:
                                        // Xử lý trường hợp không xác định
                                        break;
                                }
                            }

                        }
                    }
                    //MessageBox.Show("Dữ liệu đã được tải từ " + filePath);
                }
                else
                {
                    MessageBox.Show("File does not exist");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
            //connection
            try
            {
                string server = tbox_server.Text;
                string database = tbox_data.Text;
                string integratedSecurity = "True"; // Hoặc "False" tùy theo cài đặt
                if (combox_authentication.Text == "Window Authentication")
                {
                    UseConnectionString(server, database, integratedSecurity, "", "");
                }
                else if (combox_authentication.Text == "SQL server Authentication")
                {
                    string user = tbox_user.Text;
                    string password = tbox_pass.Text;
                    UseConnectionString(server, database, integratedSecurity, user, password);
                }

                if (sqlCon.State == ConnectionState.Open)
                {
                    // Load dữ liệu sau khi kết nối thành công
                    LoadColumnNames();
                    // Ẩn groupbox connect to server sau khi kết nối thành công
                    groupBox2.Visible = false;
                    //MessageBox.Show("Connect Successfully");
                }
                // Load dữ liệu khi UserControl được tải
                //LoadData();
                dataGridView1.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
        private void UseConnectionString(string server, string database, string integratedSecurity, string user, string password)
        {
            if (combox_authentication.Text == "Window Authentication")
            {
                strCon = $"Data Source ={server} ; Initial Catalog = {database}; Integrated Security ={integratedSecurity};";
            }
            else if (combox_authentication.Text == "SQL server Authentication")
            {
                strCon = $"Data Source={server};Initial Catalog={database};Persist Security Info={integratedSecurity};User ID={user};Password={password};";
            }
            sqlCon.ConnectionString = strCon;
            sqlCon.Open();
        }

        private void LoadColumnNames()
        {
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                {
                    MessageBox.Show("Please connect to the database before retrieving data.");
                    return;
                }

                // Lấy dữ liệu và hiển thị trong DataGridView
                int numberOfRecords = 50; // Số lượng bản ghi mới nhất muốn lấy
                string table = combox_table.Text;
                string query = $"SELECT TOP (@NumberOfRecords) id, createdAt, stopFill, startFill,net,deviation FROM {table} ORDER BY CreatedAt ASC";
                using (SqlCommand command = new SqlCommand(query, sqlCon))
                {
                    SqlParameter numberOfRecordsParameter = new SqlParameter("@NumberOfRecords", SqlDbType.Int);
                    numberOfRecordsParameter.Value = numberOfRecords;
                    command.Parameters.Add(numberOfRecordsParameter);

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        if (reader.HasRows)
                        {
                            DataTable dt = new DataTable();
                            dt.Load(reader);
                            // Thêm cột Time và Cycle Time vào DataTable
                            dt.Columns.Add("Cycle", typeof(string));
                            dt.Columns.Add("Time", typeof(string));
                            dt.Columns.Add("Cycle Time", typeof(string));
                            dt.Columns.Add("Weigh of Box", typeof(string));
                            dt.Columns.Add("Deviation", typeof(string));

                            int fetchedRowCount = 0; // Số hàng đã lấy từ cơ sở dữ liệu
                            int rowNumber = 1; // Số thứ tự của hàng
                            foreach (DataRow row in dt.Rows)
                            {
                                DateTime startFill = (DateTime)row["StartFill"];
                                DateTime stopFill = (DateTime)row["StopFill"];
                                TimeSpan cycleTime = stopFill - startFill; // Tính toán Cycle Time
                                string formattedCycleTime = string.Format("{0:D2}:{1:D2}:{2:D2}", cycleTime.Hours, cycleTime.Minutes, cycleTime.Seconds);
                                row["Cycle Time"] = formattedCycleTime;  
                                if (cycleTime.TotalMinutes >= 0 || cycleTime.TotalMinutes <= 3)
                                {
                                    // Thực hiện xử lý bổ sung tại đây nếu giá trị "Cycle Time" nằm trong khoảng [0, 3] phút
                                    row["Cycle"] = rowNumber.ToString(); // Gán số thứ tự vào cột "Cycle"
                                    row["Time"] = row["CreatedAt"].ToString();
                                    row["Weigh of Box"] = row["net"].ToString();
                                    row["Deviation"] = row["deviation"].ToString();
                                    rowNumber++; // Tăng số thứ tự hàng

                                }
                                else
                                {
                                    dt.Rows.Remove(row); // Xóa hàng nếu không thỏa điều kiện
                                }
                                fetchedRowCount++; // Tăng biến đếm số hàng đã lấy

                                if (fetchedRowCount == numberOfRecords)
                                {
                                    break; // Thoát khỏi vòng lặp khi đã lấy đủ record
                                }
                                else if (fetchedRowCount < numberOfRecords && !reader.IsClosed)
                                {
                                    if (reader.Read())
                                    {
                                        startFill = (DateTime)reader["StartFill"];
                                        stopFill = (DateTime)reader["StopFill"];
                                        cycleTime = stopFill - startFill; // Tính toán Cycle Time
                                        formattedCycleTime = string.Format("{0:D2}:{1:D2}:{2:D2}", cycleTime.Hours, cycleTime.Minutes, cycleTime.Seconds);

                                        if (cycleTime.TotalMinutes >= 0 && cycleTime.TotalMinutes <= 3)
                                        {
                                            DataRow newRow = dt.NewRow();
                                            newRow["Cycle"] = reader["id"].ToString();
                                            newRow["Time"] = reader["CreatedAt"].ToString();
                                            newRow["Cycle Time"] = formattedCycleTime;
                                            newRow["Weigh of Box"] = reader["net"].ToString();
                                            newRow["Deviation"] = reader["deviation"].ToString();
                                            dt.Rows.Add(newRow);
                                            fetchedRowCount++;
                                        }
                                    }
                                    else
                                    {
                                        reader.Close(); // Đóng reader nếu không còn dữ liệu để lấy
                                        break;
                                    }
                                }
                            }
                            // Ẩn cột StartFill và StopFill
                            dt.Columns["startFill"].ColumnMapping = MappingType.Hidden;
                            dt.Columns["stopFill"].ColumnMapping = MappingType.Hidden;
                            dt.Columns["createdAt"].ColumnMapping = MappingType.Hidden;
                            dt.Columns["id"].ColumnMapping = MappingType.Hidden;
                            dt.Columns["net"].ColumnMapping = MappingType.Hidden;
                            dt.Columns["deviation"].ColumnMapping = MappingType.Hidden;
                            dataGridView1.DataSource = dt;

                            // Đăng ký sự kiện CellFormatting
                            dataGridView1.CellFormatting += DataGridView1_CellFormatting;

                            // Thiết lập auto scale cho DataGridView
                            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                            dataGridView1.Show();
                        }
                        else
                        {
                            // Không có kết quả truy vấn
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private string CalculateCycleTime(DateTime startTime, DateTime stopTime)
        {
            TimeSpan difference = stopTime - startTime;
            string formattedTime = string.Format("{0:D2}:{1:D2}:{2:D2}", difference.Hours, difference.Minutes, difference.Seconds);
            return formattedTime;
        }

        private void DataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.ColumnIndex >= 0 && e.RowIndex >= 0)
            {
                DataGridViewColumn column = dataGridView1.Columns[e.ColumnIndex];
                if (column.ValueType == typeof(string))
                {
                    if (e.Value is DateTime dateValue)
                    {
                        e.Value = dateValue.ToString("dd/MM/yyyy hh:mm:ss tt");
                        e.FormattingApplied = true;
                    }

                }
            }
        }

        private void btn_check_Click(object sender, EventArgs e)
        {

            dataGridView1.ClearSelection();
            if (checkBox1.Checked == false)
            {
                int numberOfRecords = int.Parse(textBox2.Text); // Số lượng bản ghi mới nhất muốn lấy
                string table = combox_table.Text;
                string query = $"SELECT TOP (@NumberOfRecords) id, createdAt, stopFill, startFill,net,deviation FROM {table} ORDER BY CreatedAt ASC";
                using (SqlCommand command = new SqlCommand(query, sqlCon))
                {
                    SqlParameter numberOfRecordsParameter = new SqlParameter("@NumberOfRecords", SqlDbType.Int);
                    numberOfRecordsParameter.Value = numberOfRecords;
                    command.Parameters.Add(numberOfRecordsParameter);

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        if (reader.HasRows)
                        {
                            DataTable dt = new DataTable();
                            dt.Load(reader);

                            // Thêm cột Time và Cycle Time vào DataTable
                            dt.Columns.Add("Cycle", typeof(string));
                            dt.Columns.Add("Time", typeof(string));
                            dt.Columns.Add("Cycle Time", typeof(string));
                            dt.Columns.Add("Weigh of Box", typeof(string));
                            dt.Columns.Add("Deviation", typeof(string));

                            int fetchedRowCount = 0; // Số hàng đã lấy từ cơ sở dữ liệu
                            int rowNumber = 1; // Số thứ tự của hàng
                            foreach (DataRow row in dt.Rows)
                            {
                                DateTime startFill = (DateTime)row["StartFill"];
                                DateTime stopFill = (DateTime)row["StopFill"];
                                TimeSpan cycleTime = stopFill - startFill; // Tính toán Cycle Time
                                string formattedCycleTime = string.Format("{0:D2}:{1:D2}:{2:D2}", cycleTime.Hours, cycleTime.Minutes, cycleTime.Seconds);
                                row["Cycle Time"] = formattedCycleTime;
                                if (cycleTime.TotalMinutes >= 0 || cycleTime.TotalMinutes <= 3)
                                {
                                    // Thực hiện xử lý bổ sung tại đây nếu giá trị "Cycle Time" nằm trong khoảng [0, 3] phút
                                    row["Cycle"] = rowNumber.ToString(); // Gán số thứ tự vào cột "Cycle"
                                    row["Time"] = row["CreatedAt"].ToString();
                                    row["Weigh of Box"] = row["net"].ToString();
                                    row["Deviation"] = row["deviation"].ToString();
                                    rowNumber++; // Tăng số thứ tự hàng
                                }
                                else
                                {
                                    dt.Rows.Remove(row); // Xóa hàng nếu không thỏa điều kiện
                                }
                                fetchedRowCount++; // Tăng biến đếm số hàng đã lấy

                                if (fetchedRowCount == numberOfRecords)
                                {
                                    break; // Thoát khỏi vòng lặp khi đã lấy đủ record
                                }
                                else if (fetchedRowCount < numberOfRecords && !reader.IsClosed)
                                {
                                    if (reader.Read())
                                    {
                                        startFill = (DateTime)reader["StartFill"];
                                        stopFill = (DateTime)reader["StopFill"];
                                        cycleTime = stopFill - startFill; // Tính toán Cycle Time
                                        formattedCycleTime = string.Format("{0:D2}:{1:D2}:{2:D2}", cycleTime.Hours, cycleTime.Minutes, cycleTime.Seconds);

                                        if (cycleTime.TotalMinutes >= 0 && cycleTime.TotalMinutes <= 3)
                                        {
                                            DataRow newRow = dt.NewRow();
                                            newRow["Cycle"] = reader["id"].ToString();
                                            newRow["Time"] = reader["CreatedAt"].ToString();
                                            newRow["Cycle Time"] = formattedCycleTime;
                                            newRow["Weigh of Box"] = reader["net"].ToString();
                                            newRow["Deviation"] = reader["deviation"].ToString();
                                            dt.Rows.Add(newRow);
                                            fetchedRowCount++;
                                        }
                                    }
                                    else
                                    {
                                        reader.Close(); // Đóng reader nếu không còn dữ liệu để lấy
                                        break;
                                    }
                                }
                            }


                            // Ẩn cột StartFill và StopFill
                            dt.Columns["startFill"].ColumnMapping = MappingType.Hidden;
                            dt.Columns["stopFill"].ColumnMapping = MappingType.Hidden;
                            dt.Columns["createdAt"].ColumnMapping = MappingType.Hidden;
                            dt.Columns["id"].ColumnMapping = MappingType.Hidden;
                            dt.Columns["net"].ColumnMapping = MappingType.Hidden;
                            dt.Columns["deviation"].ColumnMapping = MappingType.Hidden;
                            dataGridView1.DataSource = dt;

                            // Đăng ký sự kiện CellFormatting
                            dataGridView1.CellFormatting += DataGridView1_CellFormatting;

                            // Thiết lập auto scale cho DataGridView
                            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                            dataGridView1.Show();
                        }
                        else
                        {
                            // Không có kết quả truy vấn
                        }
                    }
                }

            }
            else
            {

                //command
                int interval = 1;
                string table = combox_table.Text;
                string query = $"SELECT * FROM {table}  WHERE CreatedAt >= @StartDate AND CreatedAt <= @EndDate AND (DATEDIFF(DAY,  @StartDate,CreatedAt) % (@interval) = 0)";
                using (SqlCommand command = new SqlCommand(query, sqlCon))
                {
                    // Lấy giá trị thời gian từ DateTimePicker

                    DateTime startday = btn_time1.Value;

                    // Lấy ngày tiếp theo
                    DateTime nextDate = startday.AddDays(interval);
                    //lay gia tri ngay ket thuc
                    DateTime endday = btn_time2.Value;

                    //ngay bat dau
                    SqlParameter startparameter = new SqlParameter("@StartDate", SqlDbType.DateTime); // Sử dụng SqlDbType.Date để chỉ lấy thoi gian
                    SqlParameter endparameter = new SqlParameter("@EndDate", SqlDbType.DateTime); // Sử dụng SqlDbType.Date để chỉ lấy thoi gian
                    command.Parameters.AddWithValue("@Interval", interval);
                    // Thực thi truy vấn và xử lý kết quả
                    for (int i = 0; i < interval; i++)
                    {
                        if (i % interval == 0)
                        {

                            startparameter.Value = startday.AddDays(i);
                            command.Parameters.Add(startparameter);
                            //ngay ket thuc

                            endparameter.Value = endday;
                            command.Parameters.Add(endparameter);

                            //viet dieu kien 
                            // Thực thi truy vấn và xử lý kết quả

                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                if (reader.HasRows)
                                {

                                    DataTable dt = new DataTable();
                                    dt.Load(reader);
                                    // Thêm cột Time và Cycle Time vào DataTable
                                    dt.Columns.Add("Cycle", typeof(string));
                                    dt.Columns.Add("Time", typeof(string));
                                    dt.Columns.Add("Cycle Time", typeof(string));
                                    dt.Columns.Add("Weigh of Box", typeof(string));
                                    dt.Columns.Add("Deviation", typeof(string));
                                    // Đọc dữ liệu từ reader và tính toán giá trị cho cột Time và Cycle Time
                                    foreach (DataRow row in dt.Rows)
                                    {
                                        row["Cycle"] = row["id"].ToString();
                                        row["Time"] = row["CreatedAt"].ToString();
                                        row["Cycle Time"] = CalculateCycleTime((DateTime)row["StartFill"], (DateTime)row["StopFill"]);
                                        row["Weigh of Box"] = row["net"].ToString();
                                        row["Deviation"] = row["deviation"].ToString();
                                    }
                                    // Ẩn cột StartFill và StopFill
                                    dt.Columns["startFill"].ColumnMapping = MappingType.Hidden;
                                    dt.Columns["stopFill"].ColumnMapping = MappingType.Hidden;
                                    dt.Columns["createdAt"].ColumnMapping = MappingType.Hidden;
                                    dt.Columns["id"].ColumnMapping = MappingType.Hidden;
                                    dt.Columns["net"].ColumnMapping = MappingType.Hidden;
                                    dt.Columns["deviation"].ColumnMapping = MappingType.Hidden;
                                    dt.Columns["pretare"].ColumnMapping = MappingType.Hidden;
                                    dataGridView1.DataSource = dt;
                                    dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                                    dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                                    dataGridView1.Refresh();
                                }
                                else
                                {
                                    // Không có kết quả truy vấn
                                }

                            }
                            command.Parameters.Clear();
                        }


                    }
                }
            }
        }
        private System.Threading.Timer dataUpdateTimer;
        private DateTime lastEndTime;
        private IXLWorksheet ws;
        private DataTable dataGridViewStructure;
        private int exportCount = 0; // Biến đếm số lần export
        private void btn_export_Click(object sender, EventArgs e)
        {
            exportCount++; // Tăng biến đếm lên mỗi lần export
                           // Sắp xếp dữ liệu trong dataGridView1 theo cột ID
            dataGridView1.Sort(dataGridView1.Columns[0], ListSortDirection.Ascending);
            string source = tbox_source.Text;
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook workbook = excel.Workbooks.Open(source + "\\PKG15_25KG.xls");

            int newID = 1; // Biến đếm cho ID mới

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {

                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {

                    //VALUE OF CYCLE
                    if (j == 0)
                    {
                        // Thay đổi giá trị của ID bắt đầu từ 1
                        excel.Cells[i + 3, j + 1] = newID.ToString();
                        newID++; // Tăng giá trị ID mới
                    }
                    //VALUE OF DATETIME
                    if (j == 1)
                    {
                        excel.Cells[i + 3, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();

                    }
                    //VALUE OF CYCLE TIME
                    if (j == 2)
                    {
                        excel.Cells[i + 3, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();


                    }
                    //VALUE OF WEIGHT OF BOX
                    if (j == 3)
                    {
                        excel.Cells[i + 3, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();

                    }
                    //VALUE OF DEVIATION
                    if (j == 4)
                    {
                        excel.Cells[i + 3, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();

                    }
                }
            }
            excel.Visible = true;
            // Tạo tên tệp mới bao gồm ngày xuất file và số lần export
            string path = tbox_export.Text; // nhap noi xuat bao cao
            string fileName = "PKG15_25KG";
            string fileExtension = ".xls";
            string date = DateTime.Now.ToString("yyyyMMdd");
            string newFileName = $"{fileName}_{date}_Export{exportCount}{fileExtension}";
            string newFilePath = Path.Combine(path, newFileName);

            // Lưu workbook với tên và định dạng mới
            workbook.SaveAs(newFilePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
        }


        //Mode 2 che do hien thi du lieu
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                textBox2.Visible = false;
                label4.Visible = false;
            }
            else
            {
                textBox2.Visible = true;
                label4.Visible = true;
            }
        }
        private void btn_close_Click(object sender, EventArgs e)
        {
            groupBox2.Visible = false;
            dataGridView1.Visible = true;
        }

        private void btn_setting_Click(object sender, EventArgs e)
        {
            groupBox2.Visible = true;
            group_load.Visible = true;
            group_save.Visible = true;
            dataGridView1.Visible = false;
        }

        private void btn_testconnect_Click_2(object sender, EventArgs e)
        {
            if(sqlCon.State==ConnectionState.Open)
            {
                sqlCon.Close();
                if (tbox_server.Text == "")
                {
                    MessageBox.Show("Please choose server");
                }
                else
                {
                    textBox2.Text = "50";  //hiện trên textbox số record 
                    try
                    {
                        string server = tbox_server.Text;
                        string database = tbox_data.Text;
                        string integratedSecurity = "True"; // Hoặc "False" tùy theo cài đặt
                        if (combox_authentication.Text == "Window Authentication")
                        {
                            UseConnectionString(server, database, integratedSecurity, "", "");
                        }
                        else if (combox_authentication.Text == "SQL server Authentication")
                        {
                            string user = tbox_user.Text;
                            string password = tbox_pass.Text;
                            UseConnectionString(server, database, integratedSecurity, user, password);
                        }

                        if (sqlCon.State == ConnectionState.Open)
                        {
                            // Load dữ liệu sau khi kết nối thành công
                            LoadColumnNames();
                            // Ẩn groupbox connect to server sau khi kết nối thành công
                            groupBox2.Visible = false;
                            MessageBox.Show("Connect Successfully");
                        }
                        // Load dữ liệu khi UserControl được tải
                        //LoadData();
                        dataGridView1.Visible = true;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }

        }


        private void btn_save_Click(object sender, EventArgs e)
        {
            // Lấy thông tin từ các ô textbox
            string server = tbox_server.Text;
            string database = tbox_data.Text;
            string user = tbox_user.Text;
            string password = tbox_pass.Text;
            string exportPath = tbox_save.Text; // Đường dẫn tới thư mục xuất
            string export = tbox_export.Text;
            string source = tbox_source.Text;
            string table = combox_table.Text;
            string Authentication = combox_authentication.Text;
            string filePath = Path.Combine(exportPath, "test.txt");
            if (tbox_save.Text != "")
            {
                try
                {
                    if (File.Exists(exportPath))
                    {
                        File.Delete(exportPath); // Xóa tệp cũ nếu tồn tại
                    }

                    using (StreamWriter file = new StreamWriter(filePath))
                    {
                        file.WriteLine("Server:" + server);
                        file.WriteLine("User:" + user);
                        file.WriteLine("Password:" + password);
                        file.WriteLine("Database:" + database);
                        file.WriteLine("Table Name:" + table);
                        file.WriteLine("Source:" + source);
                        file.WriteLine("Export:" + export);
                        file.WriteLine("Authentication" + Authentication);

                    }

                    MessageBox.Show("File settings was created");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex.Message);
                }
            }
            else
            {
                MessageBox.Show("Please enter the path to save the file.");
            }

        }

        private void btn_loadsetting_Click(object sender, EventArgs e)
        {
            group_load.Visible = true;
            string loadsetting = tbox_load.Text;
            string filePath = Path.Combine(tbox_load.Text); // Đường dẫn từ textbox
            if (tbox_load.Text != "")
            {
                try
                {
                    if (File.Exists(filePath))
                    {
                        using (StreamReader reader = new StreamReader(filePath))
                        {
                            while (!reader.EndOfStream)
                            {
                                string line = reader.ReadLine();
                                string[] parts = line.Split(':'); // Tách chuỗi bằng dấu :

                                if (parts.Length >= 2)
                                {
                                    string field = parts[0].Trim(); // Lấy phần trước dấu :
                                    string value = string.Join(":", parts.Skip(1)).Trim(); // Lấy phần sau dấu :

                                    switch (field)
                                    {
                                        case "Server":
                                            tbox_server.Text = value;
                                            break;
                                        case "Database":
                                            tbox_data.Text = value;
                                            break;
                                        case "User":
                                            tbox_user.Text = value;
                                            break;
                                        case "Password":
                                            tbox_pass.Text = value;
                                            break;
                                        case "Source":
                                            tbox_source.Text = value;
                                            break;
                                        case "Export":
                                            tbox_export.Text = value;
                                            break;
                                        case "Table Name":
                                            combox_table.Text = value;
                                            break;
                                        default:
                                            // Xử lý trường hợp không xác định
                                            break;
                                    }
                                }

                            }
                        }
                        MessageBox.Show("The data has been loaded from " + filePath);
                    }
                    else
                    {
                        MessageBox.Show("The file does not exist.");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi: " + ex.Message);
                }
            }
            else
            {
                MessageBox.Show("Please enter the path to load the file.");
            }

        }



        private void combox_authentication_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (combox_authentication.Text == "Window Authentication")
            {
                tbox_pass.Enabled = false;
                tbox_user.Enabled = false;
                tbox_pass.BackColor = SystemColors.ControlDark; // Change to a dark color
                tbox_user.BackColor = SystemColors.ControlDark; // Change to a dark color
            }
            else if (combox_authentication.Text == "SQL server Authentication")
            {
                tbox_pass.Enabled = true;
                tbox_user.Enabled = true;
                tbox_pass.BackColor = SystemColors.Window; // Change to a light color
                tbox_user.BackColor = SystemColors.Window; // Change to a light color
            }
        }
    }
}
