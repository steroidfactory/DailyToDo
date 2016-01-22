﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Threading;

namespace ToDo
{
    public partial class Form1 : Form
    {
        delegate void SetDataTableCallback();
        private delegate void EventHandle();
        Database dbWipedrive = new Database();
        DataTable finalTable = new DataTable();
        DataTable original = new DataTable();
        public Form1()
        {
            InitializeComponent();
            dbWipedrive.initDB();
            initFinalTable();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //checkDB("5861292");
        }

        private void initFinalTable()
        {
            finalTable.Columns.Add("Serial", typeof(string));
            finalTable.Columns.Add("Status", typeof(string));
            finalTable.Columns.Add("Family", typeof(string));
            finalTable.Columns.Add("Lot", typeof(string));
            finalTable.Columns.Add("Manufacturer", typeof(string));
            finalTable.Columns.Add("Model", typeof(string));
        }


        private string checkDB(string barcode)
        {
            DateTime timeStart = new DateTime();
            DateTime timeEnd = new DateTime();
            string Name = "";
            if (dbWipedrive.Count(barcode) > 0)
            {
                timeStart = DateTime.Parse(dbWipedrive.Select(barcode)[1][0]);
                timeEnd = DateTime.Parse(dbWipedrive.Select(barcode)[2][0]);

                //Confirmed or awaiting Confirmation
                if (timeEnd > timeStart)
                {
                    Name = dbWipedrive.SelectConfirmedBy(dbWipedrive.selectOperationId(barcode)[0][0]);
                    if (Name == "")
                    {
                        return "Awaiting Confirmation";
                    }
                    else
                    {
                        return "Confirmed";
                    }
                }

                //Wiping
                else
                {
                    return "Wiping";

                }
            }
            else
            {
                return "Not started";
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

        private DataTable ReadExcelFile(string location)
        {
            DataTable dt = new DataTable();

            string connectionString = GetConnectionString(location);

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

                    //if (!sheetName.EndsWith("$"))
                    //continue;

                    // Get all rows from the Sheet
                    cmd.CommandText = "SELECT * FROM [" + sheetName + "]";

                    dt.TableName = sheetName;

                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(dt);
                }

                cmd = null;
                conn.Close();
            }
            return dt;
        }


        private void button4_Click(object sender, EventArgs e)
        {
            DialogResult result = importFile.ShowDialog();
            if(result == DialogResult.OK)
            {
                string fileName = importFile.FileName;
                //reportReadBox.Rows.Clear();
                DataTable excelData = ReadExcelFile(fileName);
                finalTable.Rows.Clear();
                barProgress.Maximum = excelData.Rows.Count - 2;
                barProgress.Value = 0;
                original = excelData;
                StartTheThread();
            }
            
        }

        /// <summary>
        /// Opens a new thread and connects to device
        /// </summary>
        public void StartTheThread()
        {
                var t = new Thread(() => addRow());
                t.Start();
        }


        private void addRow()
        {
           

            for (int i = 1; i < original.Rows.Count - 1; i++)
            {
                finalTable.Rows.Add(
                    //Barcode
                    original.Rows[i][0],
                    //Status
                    checkDB(original.Rows[i][0].ToString()),
                    //Product Family
                    original.Rows[i][1],
                    //Lot number
                    original.Rows[i][2],
                    //Manufacturer
                    original.Rows[i][4],
                    //Manufacturer Model
                    original.Rows[i][5]
                    //Started By
                    //dbWipedrive.SelectConfirmedBy(dbWipedrive.selectOperationId(excelData.Rows[i][0].ToString())[0][0])
                    );
                updateProgress();
            }
            addToTable();



        }

        private void updateProgress()
        {
            if (barProgress.InvokeRequired)
            {
                barProgress.Invoke(new Action(updateProgress));
            }
            else
            {
                barProgress.Value++;
                lblBarProgress.Text = barProgress.Value + " / " + barProgress.Maximum;
            }
        }

        private void addToTable()
        {

            // InvokeRequired required compares the thread ID of the
            // calling thread to the thread ID of the creating thread.
            // If these threads are different, it returns true.
            if (reportReadBox.InvokeRequired)
            {
                SetDataTableCallback tb = new SetDataTableCallback(addToTable);
                this.Invoke(tb, new object[] { });
            }
            else
            {
                foreach (DataRow row in finalTable.Rows)
                {
                    reportReadBox.Rows.Add(row.Field<string>(0), row.Field<string>(1), row.Field<string>(2), row.Field<string>(3), row.Field<string>(4), row.Field<string>(5));
                }
            }

        }


        private void exportFile_FileOk(object sender, CancelEventArgs e)
        {
           

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

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            exportFile.Filter = "Excel File | *.xlsx";
            DialogResult result = exportFile.ShowDialog();
            //string [] file = Directory.GetFiles(SelectedPath);
            Console.WriteLine(exportFile.FileName);
            if (exportFile.FileName != "")
            {
                reportExport(exportFile.FileName.ToString());
                exportFile.FileName = "";

            }
            else
            {
                MessageBox.Show("Please enter a file name");
            }
        }
    }
}
