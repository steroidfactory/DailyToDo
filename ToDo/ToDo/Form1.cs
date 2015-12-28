using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

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
                MessageBox.Show("No result found");
            }
        }
    }
}
