﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MySql.Data.MySqlClient;

namespace ToDo
{
    class Database
    {
        private MySqlConnection connection;
        private string server;
        private string database;
        private string uid;
        private string password;
        private string table;
        private string tableConfirmedBy;
        private string connectionSTR;


        public void initDB()
        {
            setDB();
        }

        //
        //  Connect to Database
        //
        private void setDB()
        {
            server = "10.11.3.3";
            database = "wipedrive";
            uid = "steroid";
            password = "andriyk";
            string connectionString;
            connectionString = "SERVER=" + server + ";" + "DATABASE=" +
            database + ";" + "UID=" + uid + ";" + "PASSWORD=" + password + ";" + "Convert Zero Datetime=True;";
            connection = new MySqlConnection(connectionString);
            connectionSTR = connectionString;
            table= "diskoperationlog";
            tableConfirmedBy = "customfields";

        }

        //Username - started
        //StartTime
        //EndTime
        //CustomComputerId - barcode

        //open connection to database
        private bool OpenConnection()
        {
            try
            {
                if (connection.State == System.Data.ConnectionState.Closed)
                {
                    //Console.WriteLine("Connect is to open");
                    connection.Open();
                    return true;
                }
                else if (connection.State == System.Data.ConnectionState.Open)
                {
                    //Console.WriteLine("Connection is already open");
                    return true;
                }
                if (connection.State == System.Data.ConnectionState.Closed)
                {
                    return false;
                }
                return true;

            }
            catch
            {
                //When handling errors, you can your application's response based 
                //on the error number.
                //The two most common error numbers when connecting are as follows:
                //0: Cannot connect to server.
                //1045: Invalid user name and/or password.
                //Console.WriteLine("Error in open connection is : " + ex);
                return false;

            }
        }


        //Close connection
        private bool CloseConnection()
        {
            //Console.WriteLine("Closing connectiion");
            try
            {
                connection.Close();
                return true;
            }
            catch
            {
                //Console.WriteLine("Error in close connection: " + ex);
                //MessageBox.Show(ex.Message);
                return false;
            }
        }

        //Select statement
        public List<string>[] Select(string barcode)
        {

            string query = "SELECT * FROM " + table + " WHERE CustomComputerId=" + "'" +barcode + "'" + " ORDER BY EndTime DESC" + ";";
            //int countNum = Count();
            //Create a list to store the result
            List<string>[] list = new List<string>[3];
            //for (int i = 0; i<=4; i++)
            //{
            list[0] = new List<string>();
            list[1] = new List<string>();
            list[2] = new List<string>();
            //}
            //Open connection
            if (OpenConnection() == true)
            {

                //Create Command
                MySqlCommand cmd = new MySqlCommand(query, connection);
                //Create a data reader and Execute the command
                // MySqlDataReader dataReader = cmd.ExecuteReader();

                //Read the data and store them in the list
                using (MySqlDataReader dataReader = cmd.ExecuteReader())
                {
                    while (dataReader.Read())
                    {
                        list[0].Add(dataReader["Username"] + "");
                        list[1].Add(dataReader["StartTime"] + "");
                        list[2].Add(dataReader["EndTime"] + "");

                        //Console.WriteLine(dataReader["Index_ID"].ToString() + dataReader["QT"] + dataReader["OrderNumber"]
                        //+ dataReader["ID"] + dataReader["TrackingNumber"] + dataReader["TimeIn"] );
                    }
                    //close Data Reader
                    dataReader.Close();

                    //close Connection
                    CloseConnection();
                }
                //return list to be displayed
                return list;
            }
            else
            {
                return list;
            }
        }

        //Select statement
        public int SelectCount()
        {
            string query = "SELECT * FROM " + tableConfirmedBy + " ORDER BY Id DESC LIMIT 1" + ";";
            //int countNum = Count();
            //Create a list to store the result
            int idNumber = new int();
            //for (int i = 0; i<=4; i++)
            //{
            //}

            //Open connection
            if (OpenConnection() == true)
            {

                //Create Command
                MySqlCommand cmd = new MySqlCommand(query, connection);
                //Create a data reader and Execute the command
                // MySqlDataReader dataReader = cmd.ExecuteReader();

                //Read the data and store them in the list
                using (MySqlDataReader dataReader = cmd.ExecuteReader())
                {
                    while (dataReader.Read())
                    {
                        idNumber = Int32.Parse(dataReader["Id"] + "");

                        //Console.WriteLine(dataReader["Index_ID"].ToString() + dataReader["QT"] + dataReader["OrderNumber"]
                        //+ dataReader["ID"] + dataReader["TrackingNumber"] + dataReader["TimeIn"] );
                    }
                    //close Data Reader
                    dataReader.Close();

                    //close Connection
                    CloseConnection();
                }
                //return list to be displayed
                return idNumber;
            }
            else
            {
                return idNumber;
            }
        }


        //Select statement
        public List<string>[] loadDB()
        {
            string query = "SELECT CustomComputerId, Username, StartTime, EndTime, HardDiskId FROM " +
                "(SELECT CustomComputerId, Username, StartTime, EndTime, HardDiskId FROM " + table + " WHERE DATE(StartTime) > date_sub(current_date, interval 2 month) ORDER BY EndTime DESC) tmp GROUP BY CustomComputerId" + ";";
            //int countNum = Count();
            //Create a list to store the result
            List<string>[] list = new List<string>[5];
            //for (int i = 0; i<=4; i++)
            //{
            list[0] = new List<string>();
            list[1] = new List<string>();
            list[2] = new List<string>();
            list[3] = new List<string>();
            list[4] = new List<string>();
            //}

            //Open connection
            using (MySqlConnection con = new MySqlConnection(connectionSTR))
            {
                con.Open();
                //Create Command
                MySqlCommand cmd = new MySqlCommand(query, con);
                //Create a data reader and Execute the command
                // MySqlDataReader dataReader = cmd.ExecuteReader();

                //Read the data and store them in the list
                using (MySqlDataReader dataReader = cmd.ExecuteReader())
                {
                    while (dataReader.Read())
                    {
                        list[0].Add(dataReader["CustomComputerId"] + "");
                        list[1].Add(dataReader["Username"] + "");
                        list[2].Add(dataReader["StartTime"] + "");
                        list[3].Add(dataReader["EndTime"] + "");
                        list[4].Add(dataReader["HardDiskId"] + "");
                        //Console.WriteLine(dataReader["Index_ID"].ToString() + dataReader["QT"] + dataReader["OrderNumber"]
                        //+ dataReader["ID"] + dataReader["TrackingNumber"] + dataReader["TimeIn"] );
                    }
                    //close Data Reader
                    dataReader.Close();

                    //close Connection
                    con.Close();
                }
                //return list to be displayed
                return list;
            }
        }

        //Select statement
        public List<string>[] loadConfirmation(int lastID)
        {
            string query = "SELECT DISTINCT(DiskOperationLogId), User1, UserName1, User2 FROM " + tableConfirmedBy + " WHERE Id > " + (lastID - 10000) + ";";
            //int countNum = Count();
            //Create a list to store the result
            List<string>[] list = new List<string>[4];
            //for (int i = 0; i<=4; i++)
            //{
            list[0] = new List<string>();
            list[1] = new List<string>();
            list[2] = new List<string>();
            list[3] = new List<string>();
            //}

            //Open connection
            using (MySqlConnection con = new MySqlConnection(connectionSTR))
            {
                con.Open();
                //Create Command
                MySqlCommand cmd = new MySqlCommand(query, con);
                //Create a data reader and Execute the command
                // MySqlDataReader dataReader = cmd.ExecuteReader();

                //Read the data and store them in the list
                using (MySqlDataReader dataReader = cmd.ExecuteReader())
                {
                    while (dataReader.Read())
                    {
                        list[0].Add(dataReader["DiskOperationLogId"] + "");
                        list[1].Add(dataReader["User1"] + "");
                        list[2].Add(dataReader["UserName1"] + "");
                        list[3].Add(dataReader["User2"] + "");

                        //Console.WriteLine(dataReader["Index_ID"].ToString() + dataReader["QT"] + dataReader["OrderNumber"]
                        //+ dataReader["ID"] + dataReader["TrackingNumber"] + dataReader["TimeIn"] );
                    }
                    //close Data Reader
                    dataReader.Close();

                    //close Connection
                    con.Close();
                }

                //return list to be displayed
                return list;
            }
            
        }

        //Count statement
        public int Count(string barcode)
        {
            string query = "SELECT Count(*) FROM " + table + " WHERE CustomComputerId=" + "'" + barcode + "'" + ";";
            int Count = -1;

            //Open Connection
            if (OpenConnection() == true)
            {
                //Create Mysql Command
                MySqlCommand cmd = new MySqlCommand(query, connection);

                //ExecuteScalar will return one value
                Count = int.Parse(cmd.ExecuteScalar() + "");

                //close Connection
                CloseConnection();

                return Count;
            }
            else
            {
                return Count;
            }
        }

        public string SelectConfirmedBy(string OperationId)
        {
            string query = "SELECT * FROM " + tableConfirmedBy + " WHERE DiskOperationLogId=" + "'" + OperationId + "'" + ";";
            //int countNum = Count();
            //Create a list to store the result
            string Name = "";

            //Open connection
            if (OpenConnection() == true)
            {

                //Create Command
                MySqlCommand cmd = new MySqlCommand(query, connection);
                //Create a data reader and Execute the command
                // MySqlDataReader dataReader = cmd.ExecuteReader();

                //Read the data and store them in the list
                using (MySqlDataReader dataReader = cmd.ExecuteReader())
                {
                    while (dataReader.Read())
                    {
                        Name = dataReader["User1"].ToString();
                        //Console.WriteLine(dataReader["Index_ID"].ToString() + dataReader["QT"] + dataReader["OrderNumber"]
                        //+ dataReader["ID"] + dataReader["TrackingNumber"] + dataReader["TimeIn"] );
                    }
                    //close Data Reader
                    dataReader.Close();

                    //close Connection
                    CloseConnection();
                }
                //return list to be displayed
                return Name;
            }
            else
            {
                return Name;
            }
        }

        public string SelectLoad(string queryCondition)
        {
            string query = "SELECT * FROM " + table + " WHERE " + queryCondition + ";";
            //int countNum = Count();
            //Create a list to store the result
            string Name = "";

            //Open connection
            if (OpenConnection() == true)
            {

                //Create Command
                MySqlCommand cmd = new MySqlCommand(query, connection);
                //Create a data reader and Execute the command
                // MySqlDataReader dataReader = cmd.ExecuteReader();

                //Read the data and store them in the list
                using (MySqlDataReader dataReader = cmd.ExecuteReader())
                {
                    while (dataReader.Read())
                    {
                       // Name = dataReader["User1"].ToString();
                        //Console.WriteLine(dataReader["Index_ID"].ToString() + dataReader["QT"] + dataReader["OrderNumber"]
                        //+ dataReader["ID"] + dataReader["TrackingNumber"] + dataReader["TimeIn"] );
                    }
                    //close Data Reader
                    dataReader.Close();

                    //close Connection
                    CloseConnection();
                }
                //return list to be displayed
                return Name;
            }
            else
            {
                return Name;
            }
        }

        //Select statement
        public List<string>[] selectOperationId(string barcode)
        {
            string query = "SELECT * FROM " + table + " WHERE CustomComputerId=" + "'" + barcode + "'" + ";";
            
            //int countNum = Count();
            //Create a list to store the result
            List<string>[] list = new List<string>[3];
            //for (int i = 0; i<=4; i++)
            //{
            list[0] = new List<string>();
            list[1] = new List<string>();
            list[2] = new List<string>();
            //}

            //Open connection
            if (OpenConnection() == true)
            {

                //Create Command
                MySqlCommand cmd = new MySqlCommand(query, connection);
                //Create a data reader and Execute the command
                // MySqlDataReader dataReader = cmd.ExecuteReader();

                //Read the data and store them in the list
                using (MySqlDataReader dataReader = cmd.ExecuteReader())
                {
                    while (dataReader.Read())
                    {
                        list[0].Add(dataReader["Id"] + "");
                        //Console.WriteLine(dataReader["Index_ID"].ToString() + dataReader["QT"] + dataReader["OrderNumber"]
                        //+ dataReader["ID"] + dataReader["TrackingNumber"] + dataReader["TimeIn"] );
                    }
                    //close Data Reader
                    dataReader.Close();

                    //close Connection
                    CloseConnection();
                }
                //return list to be displayed
                return list;
            }
            else
            {
                return list;
            }
        }
    }
}

