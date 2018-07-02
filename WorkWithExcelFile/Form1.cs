using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using ExcelDataReader;
using System.Reflection;
using System.Data.OleDb;


namespace WorkWithExcelFile
{
    public partial class Form1 : Form
    {
       
        public Form1()
        {
            InitializeComponent();
        }

       
       
        private void btn_Submit_Click(object sender, EventArgs e)
        {
            ///Show Excel File
            
            string connectionString = @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = c:\MainExample.xlsx; Extended Properties = 'Excel 12.0 Xml;HDR=YES'";

            using (OleDbConnection ole = new OleDbConnection(connectionString))
            {
                try
                {
                    ole.Open();

                    string sheetName = "Sheet1";

                    using (OleDbCommand oleCommand = new OleDbCommand("SELECT * FROM[" + sheetName + "$]", ole))
                    {
                        using (OleDbDataReader reader = oleCommand.ExecuteReader())
                        {
                            DataTable dataTable = new DataTable("simple");

                            dataTable.Load(reader);

                            dataGridView1.DataSource = dataTable;
                        }
                    }
                }
                catch (OleDbException oleException)
                {
                    MessageBox.Show(oleException.Message);
                }
            }
            
            
            ///End Showing Excel File


            ///Failers)

            //using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "Excel Workbook 2010|*.xlsx", ValidateNames = true })
            //{
            //    if (ofd.ShowDialog() == DialogResult.OK)
            //    {
            //        FileStream fs = File.Open(ofd.FileName, FileMode.Open, FileAccess.Read);

            //        IExcelDataReader reader = ExcelReaderFactory.CreateBinaryReader(fs);

            //        reader.GetData();




            //    }
            //}


            ///End Failers)

        }

        private void btn_Open_Click(object sender, EventArgs e)
        {
            ///sql connection and Viewed Tables on Sql
            
            string connectionString = @"Data Source = .; Initial Catalog = LoginDb; Integrated Security = True";

            using (SqlConnection sqlCon = new SqlConnection(connectionString))
            {
                sqlCon.Open();

                SqlDataAdapter sqlData = new SqlDataAdapter("Select * FROM Users", sqlCon);

                DataTable dtbl = new DataTable();

                sqlData.Fill(dtbl);

                dataGridView1.DataSource = dtbl;


            }

            /// end of Sql Showing ///

        }
    }
}
