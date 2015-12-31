using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace ToDo
{
    public partial class Form1 : Form
    {
        Database dbWipedrive = new Database();
        public Form1()
        {
            InitializeComponent();
            dbWipedrive.initDB();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            checkDB("5861292");
        }
        


        private void checkDB(string barcode)
        {
            DateTime timeStart = new DateTime();
            DateTime timeEnd = new DateTime();
            string Name = "";
            if (dbWipedrive.Count(barcode) > 0)
            {
                Console.WriteLine("Started By: " + dbWipedrive.Select(barcode)[0][0]);
                timeStart = DateTime.Parse(dbWipedrive.Select(barcode)[1][0]);
                timeEnd = DateTime.Parse(dbWipedrive.Select(barcode)[2][0]);

                //Confirmed or awaiting Confirmation
                if (timeEnd > timeStart)
                {
                    Name = dbWipedrive.SelectConfirmedBy(dbWipedrive.selectOperationId(barcode)[0][0]);
                    if (Name == "")
                    {
                        MessageBox.Show("Awaiting Confirmation");
                    }
                    else
                    {
                        MessageBox.Show("Confirmed by: " + Name);
                    }
                }

                //Wiping
                else
                {
                    MessageBox.Show("Wiping");

                }
            }
            else
            {
                MessageBox.Show("Not started");
            }
        }

        private string GetConnectionString(string location)
        {
            Dictionary<string, string> props = new Dictionary<string, string>();

            // XLSX - Excel 2007, 2010, 2012, 2013
            props["Provider"] = "Microsoft.ACE.OLEDB.12.0;";
            props["Extended Properties"] = "Excel 12.0 XML";
            props["Data Source"] = location;

            // XLS - Excel 2003 and Older
            //props["Provider"] = "Microsoft.Jet.OLEDB.4.0";
            //props["Extended Properties"] = "Excel 8.0";
            //props["Data Source"] = "C:\\MyExcel.xls";

            StringBuilder sb = new StringBuilder();

            foreach (KeyValuePair<string, string> prop in props)
            {
                sb.Append(prop.Key);
                sb.Append('=');
                sb.Append(prop.Value);
                sb.Append(';');
            }

            return sb.ToString();
        }

        private DataSet ReadExcelFile()
        {
            DataSet ds = new DataSet();

            string connectionString = GetConnectionString("C:\\excel.xlsb");

            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = conn;

                // Get all Sheets in Excel File
                DataTable dtSheet = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                // Loop through all Sheets to get data
                foreach (DataRow dr in dtSheet.Rows)
                {
                    string sheetName = dr["TABLE_NAME"].ToString();

                    if (sheetName.Contains("Lot Handling"))
                        continue;

                    if (sheetName.EndsWith("$"))
                    Console.WriteLine();

                    //if (!sheetName.EndsWith("$"))
                    //continue;

                    // Get all rows from the Sheet
                    cmd.CommandText = "SELECT * FROM [" + sheetName + "]";

                    DataTable dt = new DataTable();
                    dt.TableName = sheetName;

                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(dt);
                    ds.Tables.Add(dt);
                    Console.WriteLine(dt.Rows.Count);
                }

                cmd = null;
                conn.Close();
            }
            Console.WriteLine("returning");
            return ds;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ReadExcelFile();
        }

        private void exportFile_FileOk(object sender, CancelEventArgs e)
        {
            exportFile.Filter = "Excel File | *.xlsx";
            DialogResult result = exportFile.ShowDialog();
            // string [] file = Directory.GetFiles(SSelectedPath);
            Console.WriteLine(exportFile.FileName);
            if (exportFile.FileName != "")
            {
                //reportExport(exportSave.FileName.ToString());
                exportFile.FileName = "";

            }
            else
            {
                MessageBox.Show("Please enter a file name");
            }

        }

        private void reportExport(string location)
        {
            try
            {
                DataTable dTable = dgvToTable(reportReadBox, "dataTable");
                ExportToExcel.CreateExcelFile.CreateExcelDocument(dTable, location);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        //Converts the DataGridView to DataTable
        public DataTable dgvToTable(DataGridView dgv, String tblName)
        {
            int minRow = 0;
            DataTable dt = new DataTable(tblName);

            // Header columns
            foreach (DataGridViewColumn column in dgv.Columns)
            {
                DataColumn dc = new DataColumn(column.Name.ToString());
                dt.Columns.Add(dc);
            }

            // Data cells
            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                DataGridViewRow row = dgv.Rows[i];
                DataRow dr = dt.NewRow();
                for (int j = 0; j < dgv.Columns.Count; j++)
                {
                    dr[j] = (row.Cells[j].Value == null) ? "" : row.Cells[j].Value.ToString();
                }
                dt.Rows.Add(dr);
            }

            // Related to the bug arround min size when using ExcelLibrary for export
            for (int i = dgv.Rows.Count; i < minRow; i++)
            {
                DataRow dr = dt.NewRow();
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    dr[j] = "  ";
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }
    }
}
