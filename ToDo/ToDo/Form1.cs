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
            if (dbWipedrive.Count(barcode) > 0)
            {
                Console.WriteLine("Started By: " + dbWipedrive.Select(barcode)[0][0]);
                timeStart = DateTime.Parse(dbWipedrive.Select(barcode)[1][0]);
                timeEnd = DateTime.Parse(dbWipedrive.Select(barcode)[2][0]);
                if (timeEnd > timeStart)
                {
                    MessageBox.Show("Ended");
                }
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

        private string GetConnectionString()
        {
            Dictionary<string, string> props = new Dictionary<string, string>();

            // XLSX - Excel 2007, 2010, 2012, 2013
            props["Provider"] = "Microsoft.ACE.OLEDB.12.0;";
            props["Extended Properties"] = "Excel 12.0 XML";
            props["Data Source"] = "C:\\MyExcel.xlsx";

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
    }
}
