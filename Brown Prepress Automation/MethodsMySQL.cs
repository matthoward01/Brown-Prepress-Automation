using System;
using System.Collections.Generic;
using System.Text;
using MySql.Data.MySqlClient;
using System.Windows.Forms;

namespace Brown_Prepress_Automation
{
    class MethodsMySQL
    {
        private MySqlConnection connection;
        private string server;
        private string database;
        private string uid;
        private string password;
        private string port;

        //Constructor
        public void DBConnect()
        {
            Initialize();
            //OpenConnection();
        }

        //Initialize values
        private void Initialize()
        {
            database = "prepress";
            server = "localhost";
            //uid = "bpa";
            uid = "bpaLocal";
            if (System.IO.Directory.GetCurrentDirectory().ToLower().Contains("bin"))
            {
                database = "prepress_test";               
            }
            password = "Brown2012";
            port = "3306";
            string connectionString;
            connectionString = "SERVER=" + server + ";" + "DATABASE=" +
            database + ";" + "PORT=" +
            port + ";" + "UID=" + uid + ";" + "PASSWORD=" + password + ";";
            //database + ";" + 
            //"UID=" + uid + ";" + "PASSWORD=" + password + ";";

            connection = new MySqlConnection(connectionString);
        }        

        //open connection to database
        private bool OpenConnection()
        {
            try
            {
                connection.Open();
                return true;
            }
            catch (MySqlException ex)
            {                
                //0: Cannot connect to server.
                //1045: Invalid user name and/or password.
                switch (ex.Number)
                {
                    case 0:
                        MessageBox.Show("Cannot connect to server.  Contact administrator");
                        break;

                    case 1045:
                        MessageBox.Show("Invalid username/password, please try again");
                        break;
                }
                return false;
            }
        }

        //Close connection
        private bool CloseConnection()
        {
            try
            {
                connection.Close();
                return true;
            }
            catch (MySqlException ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
        }

        //Insert statement
        public void InsertPrepressLogAutomation(string partNumber)
        {
            string csr = "Brown Automation";
            string dateIn = DateTime.Now.ToString("M-d-yyyy");
            string timeIn = DateTime.Now.ToString("HH:mm");
            string dateOut = DateTime.Now.ToString("M-d-yyyy");
            string timeOut = DateTime.Now.ToString("HH:mm");
            string startDay = DateTime.Now.ToString("ddd");
            Int32 startTime = (Int32)(DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1))).TotalSeconds;
            Int32 endTime = (Int32)(DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1))).TotalSeconds;

            string query = "INSERT INTO prepresslog (partNumber, userIn, dateIn, timeIn, dateOut, timeOut, constructionProof, proofComplete, csr, startDay, startTime, endTime) VALUES ('" + partNumber + "', 'BAP', '" + dateIn + "', '" + timeIn + "', '" + dateOut + "', '" + timeOut + "', 'X',  'X', '" + csr + "',  '" + startDay + "',  '" + startTime.ToString() + "',  '" + endTime.ToString() + "')";
            DBConnect();
            //open connection
            if (OpenConnection() == true)
            {
                //create command and assign the query and connection from the constructor
                MySqlCommand cmd = new MySqlCommand(query, connection);

                //Execute command
                cmd.ExecuteNonQuery();

                //close connection
                CloseConnection();
            }
        }

        public void InsertPrepressLog(string partNumber, string csr, string pdfProof, string constructionProof, string specs, string additionalComments)
        {
            string dateIn = DateTime.Now.ToString("M-d-yyyy");
            string timeIn = DateTime.Now.ToString("HH:mm");
            string startDay = DateTime.Now.ToString("ddd");
            Int32 startTime = (Int32)(DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1))).TotalSeconds;

            string query = "INSERT INTO prepresslog (partNumber, userIn, dateIn, timeIn, pdfProof, constructionProof, csr, startDay, startTime, specs, additionalComments) VALUES ('" + partNumber + "', 'BAP', '" + dateIn + "', '" + timeIn + "', '" + pdfProof + "', '" + constructionProof + "', '" + csr + "',  '" + startDay + "',  '" + startTime.ToString() + "', '" + specs + "', '" + additionalComments + "')";
            DBConnect();
            //open connection
            if (OpenConnection() == true)
            {
                //create command and assign the query and connection from the constructor
                MySqlCommand cmd = new MySqlCommand(query, connection);

                //Execute command
                cmd.ExecuteNonQuery();

                //close connection
                CloseConnection();
            }
        }

        public void InsertOrders(string customer, string orderName, string fileName, string partNumber, string size, string qty, string woNumber, string soNumber, string specs)
        {

            fileName = fileName.Replace("'", "");
            orderName = orderName.Replace("'", "");
            string date = DateTime.Now.ToString("M-d-yyyy");
            string time = DateTime.Now.ToString("HH:mm");

            int orderId = SelectOrderId(orderName, date);
            string query = "";
            if (orderId == 0)
            {
                query = "INSERT INTO orders (customer, orderName, date, time) VALUES ('" + customer + "', '" + orderName + "', '" + date + "', '" + time + "')";
                DBConnect();
                //open connection
                if (OpenConnection() == true)
                {
                    //create command and assign the query and connection from the constructor
                    MySqlCommand cmd = new MySqlCommand(query, connection);

                    //Execute command
                    cmd.ExecuteNonQuery();

                    //close connection
                    CloseConnection();
                }
            }


            orderId = SelectOrderId(orderName, date);
            query = "INSERT INTO orderdetails (orderId, customer, orderName, date, time, fileName, partNumber, size, qty, woNumber, soNumber, specs) VALUES ('" + orderId + "','" + customer + "', '" + orderName + "', '" + date + "', '" + time + "', '" + fileName + "', '" + partNumber + "', '" + size + "',  '" + qty + "',  '" + woNumber + "', '" + soNumber + "', '" + specs + "')";
            DBConnect();
            //open connection
            if (OpenConnection() == true)
            {
                //create command and assign the query and connection from the constructor
                MySqlCommand cmd = new MySqlCommand(query, connection);

                //Execute command
                cmd.ExecuteNonQuery();

                //close connection
                CloseConnection();
            }
        }

        //Update statement
        public void Update()
        {
            string query = "UPDATE tableinfo SET name='Joe', age='22' WHERE name='John Smith'";

            //Open connection
            if (this.OpenConnection() == true)
            {
                //create mysql command
                MySqlCommand cmd = new MySqlCommand();
                //Assign the query using CommandText
                cmd.CommandText = query;
                //Assign the connection using Connection
                cmd.Connection = connection;

                //Execute query
                cmd.ExecuteNonQuery();

                //close connection
                this.CloseConnection();
            }
        }

        //Delete statement
        public void Delete()
        {
            string query = "DELETE FROM tableinfo WHERE name='John Smith'";

            if (this.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, connection);
                cmd.ExecuteNonQuery();
                this.CloseConnection();
            }
        }

        public int SelectOrderId(string orderName, string date)
        {
            DBConnect();
            int orderId = 0;
            string query = "SELECT id, date FROM orders WHERE orderName='" + orderName + "'";

            if (this.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, connection);
                MySqlDataReader dataReader = cmd.ExecuteReader();
                StringBuilder sb = new StringBuilder();

                while (dataReader.Read())
                {
                    if (dataReader.GetString(1).ToString() == date)
                    {
                        sb.Append(dataReader.GetString(0).ToString());
                        orderId = Int32.Parse(sb.ToString());
                    }
                }

                this.CloseConnection();

                return orderId;
            }
            else
            {
                return orderId;
            }
        }

        public string SelectCSR(string filename)
        {
            DBConnect();
            string csr = "";
            string query = "SELECT csr FROM uploads WHERE file_name='" + filename + "' ORDER BY id desc limit 1";
            
            if (this.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, connection);
                MySqlDataReader dataReader = cmd.ExecuteReader();
                StringBuilder sb = new StringBuilder();
                
                while (dataReader.Read())
                {                    
                    sb.Append(dataReader.GetString(0).ToString());
                    csr = sb.ToString();
                }

                this.CloseConnection();

                return csr;
            }
            else
            {
                return csr;
            }
        }

        //Select statement
        public List <string> [] Select(string customer)
        {
            DBConnect();
            string query = "SELECT * FROM customerSettings WHERE customer='" + customer + "'";

            List<string>[] list = new List<string>[4];
            list[0] = new List<string>();
            list[1] = new List<string>();
            list[2] = new List<string>();
            list[3] = new List<string>();

            if (this.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, connection);
                MySqlDataReader dataReader = cmd.ExecuteReader();

                while (dataReader.Read())
                {
                    list[0].Add(dataReader["customer"] + "");
                    list[1].Add(dataReader["hotFolder"] + "");
                    list[2].Add(dataReader["errorFolder"] + "");
                    list[3].Add(dataReader["archiveFolder"] + "");
                }

                dataReader.Close();

                this.CloseConnection();

                return list;
            }
            else
            {
                return list;
            }
        }

        public List<string>[] AhfDB()
        {
            DBConnect();
            string query = "SELECT * FROM ahfautomation";

            List<string>[] list = new List<string>[4];
            list[0] = new List<string>();
            list[1] = new List<string>();
            list[2] = new List<string>();
            list[3] = new List<string>();

            if (this.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, connection);
                MySqlDataReader dataReader = cmd.ExecuteReader();

                while (dataReader.Read())
                {
                    list[0].Add(dataReader["fileName"] + "");
                    list[1].Add(dataReader["fileNameAlt"] + "");
                    list[2].Add(dataReader["partNumber"] + "");
                    list[3].Add(dataReader["stock"] + "");
                }

                dataReader.Close();

                this.CloseConnection();

                return list;
            }
            else
            {
                return list;
            }
        }

        //Count statement
        //public int Count()
        //{
        //}
    }
}
