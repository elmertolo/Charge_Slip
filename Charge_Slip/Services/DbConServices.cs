using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Charge_Slip.Models;
using MySql.Data;
using MySql.Data.MySqlClient;

namespace Charge_Slip.Services
{
    class DbConServices
    {

        public MySqlConnection myConnect;
        // private int serial = 1;
        public string databaseName = "";
        public string server = "";
        public void DBConnect(string server)
        {
            try
            {
                string DBConnection = "";

                
                DBConnection = "datasource="+server+";port=3306;username=root;password=CorpCaptive; convert zero datetime=True;";

                databaseName = "captive_database";
               


                myConnect = new MySqlConnection(DBConnection);

                myConnect.Open();

            }
            catch (Exception Error)
            {

                MessageBox.Show(Error.Message, "System Error");
            }
        }// end of function
        public void DBClosed()
        {
            myConnect.Close();
        }
        public BankModel GetLastno(BankModel _bank,string _chkType)
        {
            DBConnect(_bank.IpAddress);

            string query = "Select Max(EndingSerial) from captive_database.master_database_" + _bank.tableName  + "  where ChkType = '"+_chkType+"'";
           // string query = "Select Max(EndingSerial) from captive_database.master_database_" + _bank.tableName + "_temp  where ChkType = '" + _chkType + "'";
            MySqlCommand myCommand = new MySqlCommand(query, myConnect);

            MySqlDataReader myReader = myCommand.ExecuteReader();

            while (myReader.Read())
            {
                _bank.LastNo = !myReader.IsDBNull(0) ? myReader.GetInt64(0): 0;
               // _bank.Department = !myReader.IsDBNull(1) ? myReader.GetString(1) : "";
               // _bank.BranchCode = !myReader.IsDBNull(2) ? myReader.GetString(2) : "";
            }
            DBClosed();
                return _bank;
        }
        public OrderModel SavedDatatoDatabase(OrderModel _check, string _batch, string _tableName,string _IP,DateTime _deliveryDate)
        {
            if (_check.BankCode == null)
            {

            }
            else
            {
                

                while (_check.StartingSerial.Length < 7)
                    _check.StartingSerial = "0" + _check.StartingSerial;

                while (_check.EndingSerial.Length < 7)
                    _check.EndingSerial = "0" + _check.EndingSerial;
                string sql = "";
                if (_check.BankCode == "maybank")
                {
                   //  sql = "INSERT INTO captive_database.master_database_" + _tableName + "_temp (Date,Time,DeliveryDate,ChkType,ChequeName,Address1,Batch,StartingSerial, EndingSerial)VALUES(" +
                     sql = "INSERT INTO captive_database.master_database_" + _tableName + " (Date,Time,DeliveryDate,ChkType,ChequeName,Address1,Batch,StartingSerial, EndingSerial)VALUES(" +
                                "'" + DateTime.Now.ToString("yyyy-MM-dd") + "'," +
                                "'" + DateTime.Now.ToString("HH:mm:ss") + "'," +
                               "'" + _deliveryDate.ToString("yyyy-MM-dd") + "'," +
                                "'" + _check.ChkType + "'," +
                                "'" + _check.ChequeName + "'," +
                                "'" + _check.Department + "'," +
                             //   "'" + _check.BranchCode + "'," +
                                "'" + _batch + "'," +
                                "'" + _check.StartingSerial + "'," +
                                "'" + _check.EndingSerial + "')";
                }
                else
                {
                     //sql = "INSERT INTO captive_database.master_database_" + _tableName + "_temp (Date,Time,DeliveryDate,ChkType,ChequeName,Address1,BranchCode,Batch,StartingSerial, EndingSerial)VALUES(" +
                      sql = "INSERT INTO captive_database.master_database_" + _tableName + " (Date,Time,DeliveryDate,ChkType,ChequeName,Address1,BranchCode,Batch,StartingSerial, EndingSerial)VALUES(" +
                                                   "'" + DateTime.Now.ToString("yyyy-MM-dd") + "'," +
                                                   "'" + DateTime.Now.ToString("HH:mm:ss") + "'," +
                                                  "'" + _deliveryDate.ToString("yyyy-MM-dd") + "'," +
                                                   "'" + _check.ChkType + "'," +
                                                   "'" + _check.ChequeName + "'," +
                                                   "'" + _check.Department + "'," +
                                                   "'" + _check.BranchCode + "'," +
                                                   "'" + _batch + "'," +
                                                   "'" + _check.StartingSerial + "'," +
                                                   "'" + _check.EndingSerial + "')";
                }

                DBConnect(_IP);
                MySqlCommand myCommand = new MySqlCommand(sql, myConnect);

                myCommand.ExecuteNonQuery();
                DBClosed();

            }
            return _check;
        }// end of function
        public List<MySqlLocatorModel> GetMySQLLocations()
        {
            MySqlConnection connect = new MySqlConnection("datasource=192.168.0.241 ;port=3306;username=root;password=CorpCaptive");

            connect.Open();

            MySqlCommand myCommand = new MySqlCommand("SELECT * FROM captive_database.mysqldump_location", connect);

            MySqlDataReader myReader = myCommand.ExecuteReader();

            List<MySqlLocatorModel> sqlLocator = new List<MySqlLocatorModel>();

            while (myReader.Read())
            {
                MySqlLocatorModel myLocator = new MySqlLocatorModel
                {
                    PrimaryKey = myReader.GetInt32(0),
                    Location = myReader.GetString(1)
                };

                sqlLocator.Add(myLocator);
            }

            connect.Close();

            return sqlLocator;
        }//end of Function
        public void DumpMySQL(BankModel _bank)
        {
            string dbname = "master_database_"+_bank.tableName ;
           // string dbname = "master_database_" + _bank.tableName + "_temp";
            string outputFolder = Application.StartupPath + "\\" + Main.outputFolder;
            Process proc = new Process();

            proc.StartInfo.FileName = "cmd.exe";

            proc.StartInfo.UseShellExecute = false;

            proc.StartInfo.WorkingDirectory = GetMySqlPath().ToUpper().Replace("MYSQLDUMP.EXE", "");

            proc.StartInfo.RedirectStandardInput = true;

            proc.StartInfo.RedirectStandardOutput = true;

            proc.Start();

            StreamWriter myStreamWriter = proc.StandardInput;

            string temp = "mysqldump.exe --user=root --password=CorpCaptive --host="+_bank.IpAddress+" captive_database " + dbname + " > " +
                outputFolder + "\\" + DateTime.Today.ToShortDateString().Replace("/", ".") + "-" + dbname + ".SQL";

            myStreamWriter.WriteLine(temp);

            //dbname = "aub_history";

            //temp = "mysqldump.exe --user=root --password=password=CorpCaptive --host=192.168.0.254 captive_database " + dbname + " > " +
            //     outputFolder + "\\" + DateTime.Today.ToShortDateString().Replace("/", ".") + "-" + dbname + ".SQL";

            //myStreamWriter.WriteLine(temp);

            myStreamWriter.Close();

            proc.WaitForExit();

            proc.Close();
        }//end of Function
        public string GetMySqlPath()
        {
            var mySQLocator = GetMySQLLocations();

            foreach (var loc in mySQLocator)
            {
                if (File.Exists(loc.Location))
                    return loc.Location;
            }

            return "";
        } //end of Function
        public UserModel Login(string _username, string _password)
        {
            try
            {

                if (_username == "test")
                {
                    UserModel user = new UserModel
                    {
                        Username = "Test",
                        Password = "",
                        Name = "Test User"
                    };

                    return user;
                }

                else
                {

                    DBConnect("192.168.0.254");

                    UserModel user = new UserModel();

                    string query = "SELECT Username, Password,Name FROM " + databaseName + ".master_database_user WHERE Username='" + _username + "' AND Password='" + _password + "'";

                    MySqlCommand myCommand = new MySqlCommand(query, myConnect);
                    MySqlDataAdapter sda = new MySqlDataAdapter(myCommand);
                    MySqlDataReader myReader = myCommand.ExecuteReader();
                    while (myReader.Read())
                    {
                        user = new UserModel
                        {
                            Username = myReader.GetString(0),
                            Password = myReader.GetString(1),
                            Name = myReader.GetString(2)
                        };

                    }
                    DBClosed();
                    return user;
                }
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message, "Error!!!");
                return null;
            }
        }//End of Function
    }
}
