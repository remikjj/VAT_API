using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Net;
using System.IO;
using Newtonsoft.Json;
using System.Data.SQLite;


namespace Vat_API4
{
    public partial class Form1 : Form
    {
        SQLiteConnection sqlite_conn;
        DataTable dt;
        DataTable dt2;
        DataRow dr;
        SQLiteDataAdapter adapter;
        string nip;
        string fileName;
        string date;
        string conn_str;
        string conn_insert;
        object[] array2;
        string[] array3;
        //object[] array2;


        public Form1()
        {
            InitializeComponent();
            dt = new DataTable();
            dt.Clear();
            dt.Columns.Add("Name");
            dt.Columns.Add("NIP");
            dt.Columns.Add("StatusVat");
            dt.Columns.Add("Regon");
            dt.Columns.Add("KRS");
            dt.Columns.Add("Residance Address");
            dt.Columns.Add("Working Address");
            dt.Columns.Add("Registration legal date");
            dt.Columns.Add("Request date time");
        }

        public void CreateDBNow()
        {
            SQLiteConnection.CreateFile("C:\\SQLite\\NIPdb.db");

        }
        public void ConnectNow()
        {
            sqlite_conn = new SQLiteConnection("Data Source=C:\\SQLite\\NIPdb.db;Version=3;Compress=True");
            sqlite_conn.Open();
            
        }
        public void InsertNow()
        {
           
           SQLiteCommand insert_cmd = new SQLiteCommand(conn_insert, sqlite_conn);
           insert_cmd.ExecuteNonQuery();
        }
        public void CreateTable()
        {
            conn_str = "CREATE TABLE IF NOT EXISTS test" +
                  "(name text," +
                  "nip text," +
                  "status_vat text," +
                  "regon text," +
                  "krs text," +
                  "residance_address text," +
                  "working_address text," +
                  "registration_legal_date text," +
                  "request_date_time dfdsfsdg)";

            SQLiteCommand command = new SQLiteCommand(conn_str, sqlite_conn);
            command.ExecuteNonQuery();
            //sqlite_conn.Close();
        }
        public DataTable ShowNIP(string nip, string date)
        {
            
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create("https://wl-api.mf.gov.pl/api/search/nip/"+nip+"?date="+date);
            request.Accept = "application/json";


            WebResponse response = request.GetResponse();
            StreamReader reader = new StreamReader(response.GetResponseStream());

            string NIP_Json = reader.ReadToEnd();
            Root myNIP = JsonConvert.DeserializeObject<Root>(NIP_Json);

            if (myNIP.Result.Subject != null)
            {
                array2 = new object[]
                {
                    myNIP.Result.Subject.Name.Replace('\'', ' '),
                    myNIP.Result.Subject.Nip,
                    myNIP.Result.Subject.StatusVat,
                    myNIP.Result.Subject.Regon,
                    myNIP.Result.Subject.Krs,
                    myNIP.Result.Subject.ResidenceAddress,
                    myNIP.Result.Subject.WorkingAddress,
                    myNIP.Result.Subject.RegistrationLegalDate,
                    myNIP.Result.RequestDateTime
                 };

                array3 = new string[]
                {
                    "name",
                    "nip",
                    "status_vat",
                    "regon",
                    "krs",
                    "residance_address",
                    "working_address",
                    "registration_legal_date",
                    "request_date_time"
                };

                dr = dt.NewRow();
                for (int i = 0; i <= 8; i++)
                {

                    dr[i] = array2[i];

                }
                dt.Rows.Add(dr);

                conn_insert = $@"INSERT INTO test VALUES('{array2[0]}', '{array2[1]}', '{array2[2]}','{array2[3]}','{array2[4]}',
                                  '{array2[5]}','{array2[6]}','{array2[7]}', '{array2[8]}')";
                InsertNow();

            }
            return dt;
        }
        private void BtnSelect_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Excel files (*xlsx) | *xlsx";
            openFileDialog1.ShowDialog();
            fileName = openFileDialog1.FileName;
            textBox1.Text = fileName;
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                 CreateDBNow();
            }
            catch(System.IO.IOException)
            {
                MessageBox.Show("The process cannot access the file 'file path' because it is being used by another process");
            }
            ConnectNow();
            CreateTable();
            try
            {
                ReadNow();
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                MessageBox.Show("File not found");
            }

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            try
            {
                adapter = new SQLiteDataAdapter("SELECT * FROM test WHERE nip LIKE '" + textBox2.Text + "%'", sqlite_conn);
                dt2 = new DataTable();
                adapter.Fill(dt2);
                dataGridView1.DataSource = dt2;
            }
            catch
            {}
        }
        public void ReadNow()
        {
            DateTime thisDay = DateTime.Today;
            date = thisDay.ToString("yyyy-MM-dd");
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fileName);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            int numberOfRecords = xlWorksheet.Rows.Count;
            int rowNumber = xlWorksheet.UsedRange.Rows.Count;

            for (int i = 1; i <= rowNumber; i++)
            {
                for (int j = 1; j <= 1; j++)
                {
                    //if (j == 1)
                    //textBox2.Text = textBox2.Text+("\r\n");

                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                    {
                        nip = xlRange.Cells[i, j].Value2.ToString();

                        try
                        {
                            dataGridView1.DataSource = ShowNIP(nip, date);
                        }
                        catch (WebException)
                        {
                            MessageBox.Show("Invaild NIP number in Excel file at " + i + " line");
                        }
                    }
                }
            }
            xlWorkbook.Close();
        }
        
        public void WriteNow()
        {
            Excel.Application excelApp = new Excel.Application();
            if(excelApp != null)
    {
                Excel.Workbook excelWorkbook = excelApp.Workbooks.Add();
                Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelWorkbook.Sheets.Add();

                for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        excelWorksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                    }
                }

                excelApp.ActiveWorkbook.SaveAs(@"C:\SQLite\NIP_exportt.xls", Excel.XlFileFormat.xlWorkbookNormal);

                excelWorkbook.Close();
                excelApp.Quit();

                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelWorksheet);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelWorkbook);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelApp);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            WriteNow();
        }
    }
}
