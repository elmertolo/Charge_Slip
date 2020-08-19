using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Charge_Slip.Models;
using System.Data.OleDb;
using System.Data; 
using ADOX;
using System.Threading;

namespace Charge_Slip.Services
{
    class ProcessServices
    {
        string outputFolder = Application.StartupPath;
        StreamWriter file;
        private static string GenerateSpace(int _noOfSpaces)
        {
            string output = "";

            for (int x = 0; x < _noOfSpaces; x++)
            {
                output += " ";
            }

            return output;

        }//END OF FUNCTION

        private static string Seperator()
        {
            return "";
        }
        private static string ConvertToBlockText(List<OrderModel> _check, string _prodType, string _ChkType, string _batchNumber, DateTime _deliveryDate, string _preparedBy, string _fileName)

        {

            int page = 1, lineCount = 14, blockCounter = 1, blockContent = 1;
            //string date = DateTime.Now.ToString("MMM. dd, yyyy");
            bool noFooter = true;
            string countText = "";
            string output = "";

            //Sort Check List
            var sort = (from c in _check
                        orderby c.StartingSerial
                        ascending
                        select c).ToList();

            output += "\n" + GenerateSpace(8) + "Page No. " + page.ToString() + "\n" +
            GenerateSpace(8) + ///date +
            "\n";



            output += GenerateSpace(20) + _prodType + " - " + _ChkType + "\n\n" +

             GenerateSpace(5) + "Starting Batch " + _batchNumber + ", New MICR Alignment of NCDSS is 15-54 ! ! !\n\n";


            output += GenerateSpace(8) + "BLOCK RT_NO" + GenerateSpace(5) + "M ACCT_NO" + GenerateSpace(9) + "START_NO." + GenerateSpace(2) + "END_NO.\n\n";
            //Int64 checkTypeCount = 0;

            foreach (var check in sort)
            {

                string Sserial = check.StartingSerial.ToString();
                string Eserial = check.EndingSerial.ToString();
                //if (_ChkType == "Charge Slip")
                //{
                   // checkTypeCount = check.Quantity;
                    while (Sserial.Length < 7)
                        Sserial = "0" + Sserial;

                    while (Eserial.Length < 7)
                        Eserial = "0" + Eserial;
                //}



                if (blockContent == 1)
                {
                    output += "\n" + GenerateSpace(7) + "** BLOCK " + blockCounter.ToString() + "\n";
                    lineCount += 2;
                }

                if (blockContent == 5)
                {
                    blockContent = 2;

                    blockCounter++;

                    output += "\n" + GenerateSpace(7) + "** BLOCK " + blockCounter.ToString() + "\n";

                    output += GenerateSpace(12) + blockCounter.ToString() + " 000000000" + GenerateSpace(3) + "000000000000" +
                    GenerateSpace(4) + Sserial + GenerateSpace(4) + Eserial + "\n";
                }
                else
                {
                    output += GenerateSpace(12) + blockCounter.ToString() + " 000000000" + GenerateSpace(3) + "000000000000" +
                    GenerateSpace(4) + Sserial + GenerateSpace(4) + Eserial + "\n";

                    lineCount += 1;

                    blockContent++;
                }
            }
            //if (lineCount >=61 )
            //{
            if (noFooter) //ADD FOOTER
            {
                output += "\n " + _batchNumber + GenerateSpace(46) + "DLVR: " + _deliveryDate.ToString("MM-dd(ddd)") + "\n\n" +
                    " A = " + sort.Count + GenerateSpace(20) + _fileName + ".txt\n\n" +
                    countText +
                    GenerateSpace(4) + "Prepared By" + GenerateSpace(3) + ": " + frmLogIn.userName + "\t\t\t\t RECHECKED BY:\n" +
                    GenerateSpace(4) + "Updated By" + GenerateSpace(4) + ": " + frmLogIn.userName + "\n" +
                    GenerateSpace(4) + "Time Start" + GenerateSpace(4) + ": " + DateTime.Now.ToShortTimeString() + "\n" +
                    GenerateSpace(4) + "Time Finished :\n" +
                    GenerateSpace(4) + "File rcvd" + GenerateSpace(5) + ":\n";

                noFooter = false;
            }

            // output += Seperator();

            lineCount = 1;
            //}

            return output;

        }
        public void DoBlockProcess(List<OrderModel> _checks, Main _mainForm)
        {

            var listofcheck = _checks.Select(r => r.BankCode).ToList();
            foreach (string Scheck in listofcheck)
            {

                if (Scheck == "banko_nuestra")
                {

                    string packkingListPath = outputFolder + "\\" + Main.outputFolder + "\\BlockP.txt";
                    if (File.Exists(packkingListPath))
                        File.Delete(packkingListPath);
                    var checks = _checks.Where(a => a.BankCode == Scheck).Distinct().ToList();
                    file = File.CreateText(packkingListPath);
                    file.Close();

                    using (file = new StreamWriter(File.Open(packkingListPath, FileMode.Append)))
                    {

                        string output = ConvertToBlockText(checks, "Banko Nuestra", "Charge Slip", _mainForm.batchfile, _mainForm.deliveryDate, frmLogIn.userName, _mainForm.FileName);

                        file.WriteLine(output);
                    }

                }
            }
            foreach (string Scheck in listofcheck)
            {

                if (Scheck == "mbtc")
                {

                    string packkingListPath = outputFolder + "\\" + Main.outputFolder + "\\BlockP.txt";
                    if (File.Exists(packkingListPath))
                        File.Delete(packkingListPath);
                    var checks = _checks.Where(a => a.BankCode == Scheck).Distinct().ToList();
                    file = File.CreateText(packkingListPath);
                    file.Close();

                    using (file = new StreamWriter(File.Open(packkingListPath, FileMode.Append)))
                    {

                        string output = ConvertToBlockText(checks, "Metro Bank", "Charge Slip", _mainForm.batchfile, _mainForm.deliveryDate, frmLogIn.userName, _mainForm.FileName);

                        file.WriteLine(output);
                    }

                }
            }
            foreach (string Scheck in listofcheck)
            {

                if (Scheck == "union")
                {

                    string packkingListPath = outputFolder + "\\" + Main.outputFolder + "\\BlockP.txt";
                    if (File.Exists(packkingListPath))
                        File.Delete(packkingListPath);
                    var checks = _checks.Where(a => a.BankCode == Scheck).Distinct().ToList();
                    file = File.CreateText(packkingListPath);
                    file.Close();

                    using (file = new StreamWriter(File.Open(packkingListPath, FileMode.Append)))
                    {

                        string output = ConvertToBlockText(checks, "Union Bank", "Check Voucher", _mainForm.batchfile, _mainForm.deliveryDate, frmLogIn.userName, _mainForm.FileName);

                        file.WriteLine(output);
                    }

                }
            }
            foreach (string Scheck in listofcheck)
            {

                if (Scheck == "sbtc")
                {

                    string packkingListPath = outputFolder + "\\" + Main.outputFolder + "\\BlockP.txt";
                    if (File.Exists(packkingListPath))
                        File.Delete(packkingListPath);
                    var checks = _checks.Where(a => a.BankCode == Scheck).Distinct().ToList();
                    file = File.CreateText(packkingListPath);
                    file.Close();

                    using (file = new StreamWriter(File.Open(packkingListPath, FileMode.Append)))
                    {

                        string output = ConvertToBlockText(checks, "Security Bank", "Charge Slip", _mainForm.batchfile, _mainForm.deliveryDate, frmLogIn.userName, _mainForm.FileName);

                        file.WriteLine(output);
                    }

                }
            }
            foreach (string Scheck in listofcheck)
            {

                if (Scheck == "maybank")
                {

                    string packkingListPath = outputFolder + "\\" + Main.outputFolder + "\\BlockP.txt";
                    if (File.Exists(packkingListPath))
                        File.Delete(packkingListPath);
                    var checks = _checks.Where(a => a.BankCode == Scheck).Distinct().ToList();
                    file = File.CreateText(packkingListPath);
                    file.Close();

                    using (file = new StreamWriter(File.Open(packkingListPath, FileMode.Append)))
                    {

                        string output = ConvertToBlockText(checks, "Maybank", "Charge Slip", _mainForm.batchfile, _mainForm.deliveryDate, frmLogIn.userName, _mainForm.FileName);

                        file.WriteLine(output);
                    }

                }
            }
        }

        public static void CreteaMDB(List<OrderModel> _orders, string _batch, string _chkType, string _fileName)
        {
            string txtName = "";
            string fileName = "";
            //if (_chkType == "CS" || )
            //{
                fileName = Application.StartupPath + "\\MDB\\" + _fileName;
                //  txtName = "CS" + _batch.Substring(0, 4) + "C.txt";
          //  }
            //else if (_chkType == "SR")
            //{
            //    fileName = Application.StartupPath + "\\MDB\\" + _fileName;
            //    txtName = "ISL" + _batch.Substring(0, 4) + "S.20P";
            //}

            FileInfo fileInfo = new FileInfo(fileName);

            if (fileInfo.Exists)
                fileInfo.Delete();

            string mdbConn = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + fileName + "; Jet OLEDB:Engine Type=5";

            Catalog tClass = new Catalog();

            tClass.Create(mdbConn);

            System.Runtime.InteropServices.Marshal.ReleaseComObject(tClass);

            GC.Collect();

            tClass = new Catalog();

            tClass.let_ActiveConnection(mdbConn);

            for (int x = 1; x <= 4; x++)
            {

                Table tTable = new Table();

                string tableTemp = "";

                if (x == 1)
                    tableTemp = "Out";
                else
                    tableTemp = "Outs";

                tTable.Name = "InputFile_" + x.ToString() + tableTemp;

                //COLUMN NAMES
                //tTable.Columns.Append("BRSTN", DataTypeEnum.adVarWChar, 10);

                //tTable.Columns.Append("AccountNumber", DataTypeEnum.adVarWChar, 30);

                //tTable.Columns.Append("RT1to5", DataTypeEnum.adVarWChar, 5);

                //tTable.Columns.Append("RT6to9", DataTypeEnum.adVarWChar, 5);

                //tTable.Columns.Append("AccountNumberWithHypen", DataTypeEnum.adVarWChar, 30);

                //tTable.Columns.Append("Serial", DataTypeEnum.adVarWChar, 10);

                //tTable.Columns.Append("Name1", DataTypeEnum.adVarWChar, 50);

                //tTable.Columns.Append("Name2", DataTypeEnum.adVarWChar, 50);

                //tTable.Columns.Append("Name3", DataTypeEnum.adVarWChar, 50);

                //tTable.Columns.Append("Address1", DataTypeEnum.adVarWChar, 100);

                //tTable.Columns.Append("Address2", DataTypeEnum.adVarWChar, 100);

                //tTable.Columns.Append("Address3", DataTypeEnum.adVarWChar, 100);

                //tTable.Columns.Append("Address4", DataTypeEnum.adVarWChar, 100);

                //tTable.Columns.Append("Address5", DataTypeEnum.adVarWChar, 100);

                //  tTable.Columns.Append("Address6", DataTypeEnum.adVarWChar, 100);

                tTable.Columns.Append("BankName", DataTypeEnum.adVarWChar, 100);

                tTable.Columns.Append("BranchCode", DataTypeEnum.adVarWChar, 100);

                tTable.Columns.Append("Department", DataTypeEnum.adVarWChar, 100);

                tTable.Columns.Append("Serial", DataTypeEnum.adVarWChar, 10);

                tTable.Columns.Append("StartingSerial", DataTypeEnum.adVarWChar, 20);

                tTable.Columns.Append("EndingSerial", DataTypeEnum.adVarWChar, 20);

                tTable.Columns.Append("PcsPerBook", DataTypeEnum.adVarWChar, 3);

                tTable.Columns.Append("FileName", DataTypeEnum.adVarWChar, 30);

                tTable.Columns.Append("PrimaryKey", DataTypeEnum.adVarWChar);

                tTable.Columns["PrimaryKey"].Attributes = ColumnAttributesEnum.adColNullable;

                tTable.Columns.Append("PageNumber", DataTypeEnum.adVarWChar);

                tTable.Columns.Append("DataNumber", DataTypeEnum.adVarWChar, 20);

                tClass.Tables.Append((object)tTable);
            }//END FOR

            GC.Collect(); //IF NOT INCLUDED FILE REMAIN IN OPENED STATUS

            OleDbConnection connection = new OleDbConnection(mdbConn);

            OleDbCommand cmd = new OleDbCommand();

            cmd.Connection = connection;

            connection.Open();

            int primaryKey = 1;



            int dataNumber1 = 0, dataNumber2 = 0, dataNumber3 = 0, dataNumber4 = 0;//SERVE AS DATANUMBER
            #region 1Out Format
            //1OUTS FORMAT
            foreach (var check in _orders)
            {
                //  string RTFirst = check.BRSTN.Substring(0, 5);

                //  string RTLast = check.BRSTN.Substring(check.BRSTN.Length - 4, 4);

                string startSeries = check.StartingSerial.ToString();

                string endSeries = check.EndingSerial.ToString();

                //  string acctNoHypen = check.AccountNo.Substring(0, 3) + "-" + check.AccountNo.Substring(3, 6) + "-" + check.AccountNo.Substring(9, 3);

                while (startSeries.Length < 10)
                    startSeries = "0" + startSeries;

                while (endSeries.Length < 10)
                    endSeries = "0" + endSeries;

                Int64 endS =Int64.Parse(check.EndingSerial);
                Int64 start = Int64.Parse(check.StartingSerial);

                while (start <= endS)
                {
                    string temp = start.ToString();

                    while (temp.Length < 10)
                        temp = "0" + temp;

                    cmd.CommandText = "INSERT INTO InputFile_1Out (BankName,BranchCode,Department, Serial,StartingSerial, EndingSerial, " +
                        "PcsPerBook, FileName, PrimaryKey, PageNumber, DataNumber) " +
                        "VALUES ('" + check.BankName.ToUpper() + "','" + check.BranchCode + "','" + check.Department + "','" + temp + "','" + startSeries + "','" + endSeries +
                        "','50','" + txtName + "','" + primaryKey + "','0','" + primaryKey + "');";

                    cmd.ExecuteNonQuery();

                    start++;

                    primaryKey++;
                }//END WHILE
            }//END FOR
            #endregion
            #region 2Outs Format
            //2Outs Format
            dataNumber1 = 0;
            dataNumber2 = 50; //SERVE AS DATANUMBER

            primaryKey = 0;

            for (int x1 = 0; x1 < _orders.Count; x1++)
            {
                Int64 start1 = Int64.Parse(_orders[x1].StartingSerial);

                //   string RTFirst1 = _orders[x1].BRSTN.Substring(0, 5);

                // string RTLast1 = _orders[x1].BRSTN.Substring(_orders[x1].BRSTN.Length - 4, 4);

                string startSeries1 = _orders[x1].StartingSerial.ToString();

                string endSeries1 = _orders[x1].EndingSerial.ToString();

                //        string acctNoHypen1 = _orders[x1].AccountNo.Substring(0, 3) + "-" + _orders[x1].AccountNo.Substring(3, 6) + "-" + _orders[x1].AccountNo.Substring(9, 3);

                while (startSeries1.Length < 10)
                    startSeries1 = "0" + startSeries1;

                while (endSeries1.Length < 10)
                    endSeries1 = "0" + endSeries1;

                int x2 = 0;

                Int64 start2 = 0;

                string startSeries2 = "", endSeries2 = "";

                if (x1 + 1 < _orders.Count)
                {
                    x2 = x1 + 1;

                    //RTFirst2 = _orders[x2].BRSTN.Substring(0, 5);

                    //RTLast2 = _orders[x2].BRSTN.Substring(_orders[x2].BRSTN.Length - 4, 4);

                    startSeries2 = _orders[x2].StartingSerial;

                    endSeries2 = _orders[x2].EndingSerial;

                    //  acctNoHypen2 = _orders[x2].AccountNo.Substring(0, 3) + "-" + _orders[x2].AccountNo.Substring(3, 6) + "-" + _orders[x2].AccountNo.Substring(9, 3);

                    while (startSeries2.Length < 10)
                        startSeries2 = "0" + startSeries2;

                    while (endSeries2.Length < 10)
                        endSeries2 = "0" + endSeries2;

                    start2 = Int64.Parse(_orders[x2].StartingSerial);

                    x1++;
                }

                for (int x = 0; x < 50; x++)
                {
                    dataNumber1++;

                    string temp1 = start1.ToString(); //Serial with 0 format

                    while (temp1.Length < 10)
                        temp1 = "0" + temp1;

                    cmd.CommandText = "INSERT INTO InputFile_2Outs (BankName,BranchCode,Department, Serial, StartingSerial, EndingSerial, " +
                        "PcsPerBook, FileName, PrimaryKey, PageNumber, DataNumber) " +
                        "VALUES('" + _orders[x1].BankName.ToUpper() + "','" + _orders[x1].BranchCode + "','" + _orders[x1].Department + "','" + temp1 + "','" + startSeries1 + "','" + endSeries1 +
                        "','50','" + txtName + "','" + primaryKey + "','0','" + dataNumber1 + "');";

                    cmd.ExecuteNonQuery();

                    start1++;

                    primaryKey++;

                    if (x1 + 1 < _orders.Count)
                    {
                        dataNumber2++;

                        string temp2 = start2.ToString();

                        while (temp2.Length < 10)
                            temp2 = "0" + temp2;

                        cmd.CommandText = "INSERT INTO InputFile_2Outs (BankName,BranchCode,Department,Serial, StartingSerial, EndingSerial, " +
                        "PcsPerBook, FileName, PrimaryKey, PageNumber, DataNumber) " +
                        "VALUES ('" + _orders[x2].BankName.ToUpper() + "','" + _orders[x2].BranchCode + "','" + _orders[x2].Department + "','" + temp2 + "','" + startSeries2 + "','" + endSeries2 +
                        "','50','" + txtName + "','" + primaryKey + "','0','" + dataNumber2 + "');";

                        cmd.ExecuteNonQuery();

                        start2++;
                        primaryKey++;
                    }
                    else // INSERT BLANK FIELD
                    {
                        dataNumber2++;

                        cmd.CommandText = "INSERT INTO InputFile_2Outs (BankName,BranchCode,Department, Serial, StartingSerial, EndingSerial, " +
                        "PcsPerBook, FileName, PrimaryKey, PageNumber, DataNumber) " +
                        "VALUES ('','','','','','','','','" + primaryKey + "','0','" + dataNumber2 + "');";

                        cmd.ExecuteNonQuery();

                        primaryKey++;
                    }

                }//END WHILE
            }//END FOR
            #endregion
            #region 3Outs Format
            dataNumber1 = 0;
            dataNumber2 = 50; //SERVE AS DATANUMBER
            dataNumber3 = 50 * 2;

            primaryKey = 0;

            for (int x1 = 0; x1 < _orders.Count; x1++)
            {
                Int64 start1 = Int64.Parse(_orders[x1].StartingSerial);

                //  string RTFirst1 = _orders[x1].BRSTN.Substring(0, 5);

                //  string RTLast1 = _orders[x1].BRSTN.Substring(_orders[x1].BRSTN.Length - 4, 4);

                string startSeries1 = _orders[x1].StartingSerial.ToString();

                string endSeries1 = _orders[x1].EndingSerial.ToString();

                //  string acctNoHypen1 = _orders[x1].AccountNo.Substring(0, 3) + "-" + _orders[x1].AccountNo.Substring(3, 6) + "-" + _orders[x1].AccountNo.Substring(9, 3);

                while (startSeries1.Length < 10)
                    startSeries1 = "0" + startSeries1;

                while (endSeries1.Length < 10)
                    endSeries1 = "0" + endSeries1;

                bool Out2 = false, Out3 = false;

                int x2 = 0;

                Int64 start2 = 0;

                string startSeries2 = "", endSeries2 = "";

                if (x1 + 1 < _orders.Count)
                {
                    x2 = x1 + 1;

                    //RTFirst2 = _orders[x2].BRSTN.Substring(0, 5);

                    //RTLast2 = _orders[x2].BRSTN.Substring(_orders[x2].BRSTN.Length - 4, 4);

                    startSeries2 = _orders[x2].StartingSerial.ToString();

                    endSeries2 = _orders[x2].EndingSerial.ToString();

                    //   acctNoHypen2 = _orders[x2].AccountNo.Substring(0, 3) + "-" + _orders[x2].AccountNo.Substring(3, 6) + "-" + _orders[x2].AccountNo.Substring(9, 3);

                    while (startSeries2.Length < 10)
                        startSeries2 = "0" + startSeries2;

                    while (endSeries2.Length < 10)
                        endSeries2 = "0" + endSeries2;

                    start2 = Int64.Parse(_orders[x2].StartingSerial);

                    x1++;

                    Out2 = true;
                }
                else
                    Out2 = false;

                int x3 = 0;

                Int64 start3 = 0;

                string startSeries3 = "", endSeries3 = "";

                if (x1 + 1 < _orders.Count)
                {
                    x3 = x1 + 1;

                    // RTFirst3 = _orders[x3].BRSTN.Substring(0, 5);

                    // RTLast3 = _orders[x3].BRSTN.Substring(_orders[x3].BRSTN.Length - 4, 4);

                    startSeries3 = _orders[x3].StartingSerial.ToString();

                    endSeries3 = _orders[x3].EndingSerial.ToString();

                    //  acctNoHypen3 = _orders[x3].AccountNo.Substring(0, 3) + "-" + _orders[x3].AccountNo.Substring(3, 6) + "-" + _orders[x3].AccountNo.Substring(9, 3);

                    while (startSeries3.Length < 10)
                        startSeries3 = "0" + startSeries3;

                    while (endSeries3.Length < 10)
                        endSeries3 = "0" + endSeries3;

                    start3 = Int64.Parse(_orders[x3].StartingSerial);

                    x1++;

                    Out3 = true;
                }
                else
                    Out3 = false;

                for (int x = 0; x < 50; x++)
                {
                    dataNumber1++;

                    string temp1 = start1.ToString(); //Serial with 0 format

                    while (temp1.Length < 10)
                        temp1 = "0" + temp1;

                    cmd.CommandText = "INSERT INTO InputFile_3Outs (BankName,BranchCode,Department,Serial, StartingSerial, EndingSerial, " +
                        "PcsPerBook, FileName, PrimaryKey, PageNumber, DataNumber) " +
                        "VALUES ('" + _orders[x1].BankName.ToUpper() + "','" + _orders[x1].BranchCode + "','" + _orders[x1].Department + "','" + temp1 + "','" + startSeries1 + "','" + endSeries1 +
                        "','50','" + txtName + "','" + primaryKey + "','0','" + dataNumber1 + "');";

                    cmd.ExecuteNonQuery();

                    start1++;

                    primaryKey++;

                    if (Out2)
                    {
                        dataNumber2++;

                        string temp2 = start2.ToString();

                        while (temp2.Length < 10)
                            temp2 = "0" + temp2;

                        cmd.CommandText = "INSERT INTO InputFile_3Outs (BankName,BranchCode,Department,Serial, StartingSerial, EndingSerial, " +
                        "PcsPerBook, FileName, PrimaryKey, PageNumber, DataNumber) " +
                        "VALUES ('" + _orders[x2].BankName.ToUpper() + "','" + _orders[x2].BranchCode + "','" + _orders[x2].Department + "','" + temp2 + "','" + startSeries2 + "','" + endSeries2 +
                        "','50','" + txtName + "','" + primaryKey + "','0','" + dataNumber2 + "');";

                        cmd.ExecuteNonQuery();

                        start2++;
                        primaryKey++;
                    }
                    else // INSERT BLANK FIELD
                    {
                        dataNumber2++;

                        cmd.CommandText = "INSERT INTO InputFile_3Outs (BankName,BranchCode,Department,Serial, StartingSerial, EndingSerial, " +
                        "PcsPerBook, FileName, PrimaryKey, PageNumber, DataNumber) " +
                        "VALUES ('','','','','','','','','" + primaryKey + "','0','" + dataNumber2 + "');";

                        cmd.ExecuteNonQuery();

                        primaryKey++;
                    }

                    if (Out3)
                    {
                        dataNumber3++;

                        string temp3 = start3.ToString();

                        while (temp3.Length < 10)
                            temp3 = "0" + temp3;

                        cmd.CommandText = "INSERT INTO InputFile_3Outs (BankName,BranchCode,Department ,Serial,StartingSerial, EndingSerial, " +
                        "PcsPerBook, FileName, PrimaryKey, PageNumber, DataNumber) " +
                        "VALUES ('" + _orders[x3].BankName.ToUpper() + "','" + _orders[x3].BranchCode + "','" + _orders[x3].Department + "','" + temp3 + "','" + startSeries3 + "','" + endSeries3 +
                        "','50','" + txtName + "','" + primaryKey + "','0','" + dataNumber3 + "');";

                        cmd.ExecuteNonQuery();

                        start3++;
                        primaryKey++;
                    }
                    else // INSERT BLANK FIELD
                    {
                        dataNumber2++;

                        cmd.CommandText = "INSERT INTO InputFile_3Outs (BankName,BranchCode,Department,Serial, StartingSerial, EndingSerial, " +
                        "PcsPerBook, FileName, PrimaryKey, PageNumber, DataNumber) " +
                        "VALUES ('','','','','','','','','" + primaryKey + "','0','" + dataNumber2 + "');";

                        cmd.ExecuteNonQuery();

                        primaryKey++;
                    }


                }//END WHILE
            }//END FOR

            #endregion
            #region 4Outs Format
            dataNumber1 = 0;
            dataNumber2 = 50; //SERVE AS DATANUMBER
            dataNumber3 = 50 * 2;
            dataNumber4 = 50 * 3;

            primaryKey = 0;

            for (int x1 = 0; x1 < _orders.Count; x1++)
            {
                Int64 start1 = Int64.Parse(_orders[x1].StartingSerial);

                //string RTFirst1 = _orders[x1].BRSTN.Substring(0, 5);

                //string RTLast1 = _orders[x1].BRSTN.Substring(_orders[x1].BRSTN.Length - 4, 4);

                string startSeries1 = _orders[x1].StartingSerial.ToString();

                string endSeries1 = _orders[x1].EndingSerial.ToString();

                // string acctNoHypen1 = _orders[x1].AccountNo.Substring(0, 3) + "-" + _orders[x1].AccountNo.Substring(3, 6) + "-" + _orders[x1].AccountNo.Substring(9, 3);

                while (startSeries1.Length < 10)
                    startSeries1 = "0" + startSeries1;

                while (endSeries1.Length < 10)
                    endSeries1 = "0" + endSeries1;

                int x2 = 0;

                Int64 start2 = 0;

                string startSeries2 = "", endSeries2 = "";

                bool Out2 = false, Out3 = false, Out4 = false;

                //For 2ndLine
                if (x1 + 1 < _orders.Count)
                {
                    x2 = x1 + 1;

                    //   RTFirst2 = _orders[x2].BRSTN.Substring(0, 5);

                    //  RTLast2 = _orders[x2].BRSTN.Substring(_orders[x2].BRSTN.Length - 4, 4);

                    startSeries2 = _orders[x2].StartingSerial.ToString();

                    endSeries2 = _orders[x2].EndingSerial.ToString();

                    //   acctNoHypen2 = _orders[x2].AccountNo.Substring(0, 3) + "-" + _orders[x2].AccountNo.Substring(3, 6) + "-" + _orders[x2].AccountNo.Substring(9, 3);

                    while (startSeries2.Length < 10)
                        startSeries2 = "0" + startSeries2;

                    while (endSeries2.Length < 10)
                        endSeries2 = "0" + endSeries2;

                    start2 = Int64.Parse(_orders[x2].StartingSerial);

                    x1++;

                    Out2 = true;
                }
                else
                    Out2 = false;

                int x3 = 0;

                Int64 start3 = 0;

                string /*RTFirst3 = "" RTLast3 = "",*/ startSeries3 = "", endSeries3 = "";

                //FOR 3rdLine
                if (x1 + 1 < _orders.Count)
                {
                    x3 = x1 + 1;

                    //  RTFirst3 = _orders[x3].BRSTN.Substring(0, 5);

                    //   RTLast3 = _orders[x3].BRSTN.Substring(_orders[x3].BRSTN.Length - 4, 4);

                    startSeries3 = _orders[x3].StartingSerial.ToString();

                    endSeries3 = _orders[x3].EndingSerial.ToString();

                    // acctNoHypen3 = _orders[x3].AccountNo.Substring(0, 3) + "-" + _orders[x3].AccountNo.Substring(3, 6) + "-" + _orders[x3].AccountNo.Substring(9, 3);

                    while (startSeries3.Length < 10)
                        startSeries3 = "0" + startSeries3;

                    while (endSeries3.Length < 10)
                        endSeries3 = "0" + endSeries3;

                    start3 = Int64.Parse(_orders[x3].StartingSerial);

                    x1++;

                    Out3 = true;
                }
                else
                    Out3 = false;

                int x4 = 0;

                Int64 start4 = 0;

                string startSeries4 = "", endSeries4 = "";

                //For 4thLine
                if (x1 + 1 < _orders.Count)
                {
                    x4 = x1 + 1;

                    //RTFirst4 = _orders[x4].BRSTN.Substring(0, 5);

                    //RTLast4 = _orders[x4].BRSTN.Substring(_orders[x4].BRSTN.Length - 4, 4);

                    startSeries4 = _orders[x4].StartingSerial.ToString();

                    endSeries4 = _orders[x4].EndingSerial.ToString();

                    //acctNoHypen4 = _orders[x4].AccountNo.Substring(0, 3) + "-" + _orders[x4].AccountNo.Substring(3, 6) + "-" + _orders[x4].AccountNo.Substring(9, 3);

                    while (startSeries4.Length < 10)
                        startSeries4 = "0" + startSeries4;

                    while (endSeries4.Length < 10)
                        endSeries4 = "0" + endSeries4;

                    start4 = Int64.Parse(_orders[x4].StartingSerial);

                    x1++;

                    Out4 = true;
                }
                else
                    Out4 = false;

                for (int x = 0; x < 50; x++)
                {
                    dataNumber1++;

                    string temp1 = start1.ToString(); //Serial with 0 format

                    while (temp1.Length < 10)
                        temp1 = "0" + temp1;

                    cmd.CommandText = "INSERT INTO InputFile_4Outs (BankName,BranchCode,Department,Serial, StartingSerial, EndingSerial, " +
                        "PcsPerBook, FileName, PrimaryKey, PageNumber, DataNumber) " +
                        "VALUES ('" + _orders[x1].BankName.ToUpper() + "','" + _orders[x1].BranchCode + "','" + _orders[x1].Department + "','" + temp1 + "','" + startSeries1 + "','" + endSeries1 +
                        "','50','" + txtName + "','" + primaryKey + "','0','" + dataNumber1 + "');";

                    cmd.ExecuteNonQuery();

                    start1++;

                    primaryKey++;

                    if (Out2)
                    {
                        dataNumber2++;

                        string temp2 = start2.ToString();

                        while (temp2.Length < 10)
                            temp2 = "0" + temp2;

                        cmd.CommandText = "INSERT INTO InputFile_4Outs (BankName,BranchCode,Department,Serial, StartingSerial, EndingSerial, " +
                        "PcsPerBook, FileName, PrimaryKey, PageNumber, DataNumber) " +
                        "VALUES ('" + _orders[x2].BankName.ToUpper() + "','" + _orders[x2].BranchCode + "','" + _orders[x2].Department + "','" + temp2 + "','" + startSeries2 + "','" + endSeries2 +
                        "','50','" + txtName + "','" + primaryKey + "','0','" + dataNumber2 + "');";

                        cmd.ExecuteNonQuery();

                        start2++;
                        primaryKey++;
                    }
                    else // INSERT BLANK FIELD
                    {
                        dataNumber2++;

                        cmd.CommandText = "INSERT INTO InputFile_4Outs (BankName,BranchCode,Department,Serial, StartingSerial, EndingSerial, " +
                        "PcsPerBook, FileName, PrimaryKey, PageNumber, DataNumber) " +
                        "VALUES ('','','','','','','','','" + primaryKey + "','0','" + dataNumber2 + "');";

                        cmd.ExecuteNonQuery();

                        primaryKey++;
                    }

                    if (Out3)
                    {
                        dataNumber3++;

                        string temp3 = start3.ToString();

                        while (temp3.Length < 10)
                            temp3 = "0" + temp3;

                        cmd.CommandText = "INSERT INTO InputFile_4Outs (BankName,BranchCode,Department,Serial, StartingSerial, EndingSerial, " +
                        "PcsPerBook, FileName, PrimaryKey, PageNumber, DataNumber) " +
                        "VALUES ('" + _orders[x3].BankName.ToUpper() + "','" + _orders[x3].BranchCode + "','" + _orders[x3].Department + "','" + temp3 + "','" + startSeries3 + "','" + endSeries3 +
                        "','50','" + txtName + "','" + primaryKey + "','0','" + dataNumber3 + "');";

                        cmd.ExecuteNonQuery();

                        start3++;
                        primaryKey++;
                    }
                    else // INSERT BLANK FIELD
                    {
                        dataNumber2++;

                        cmd.CommandText = "INSERT INTO InputFile_4Outs (BankName,BranchCode,Department,Serial, StartingSerial, EndingSerial, " +
                        "PcsPerBook, FileName, PrimaryKey, PageNumber, DataNumber) " +
                        "VALUES ('','','','','','','','','" + primaryKey + "','0','" + dataNumber2 + "');";

                        cmd.ExecuteNonQuery();

                        primaryKey++;
                    }

                    if (Out4)
                    {
                        dataNumber4++;

                        string temp4 = start4.ToString();

                        while (temp4.Length < 10)
                            temp4 = "0" + temp4;

                        cmd.CommandText = "INSERT INTO InputFile_4Outs (BankName,BranchCode,Department,Serial, StartingSerial, EndingSerial, " +
                        "PcsPerBook, FileName, PrimaryKey, PageNumber, DataNumber) " +
                        "VALUES ('" + _orders[x4].BankName.ToUpper() + "','" + _orders[x4].BranchCode + "','" + _orders[x4].Department + "','" + temp4 + "','" + startSeries4 + "','" + endSeries4 +
                        "','50','" + txtName + "','" + primaryKey + "','0','" + dataNumber4 + "');";

                        cmd.ExecuteNonQuery();

                        start4++;

                        primaryKey++;
                    }
                    else // INSERT BLANK FIELD
                    {
                        dataNumber4++;

                        cmd.CommandText = "INSERT INTO InputFile_4Outs (BankName,BranchCode,Department,Serial, StartingSerial, EndingSerial, " +
                        "PcsPerBook, FileName, PrimaryKey, PageNumber, DataNumber) " +
                        "VALUES ('','','','','','','','','" + primaryKey + "','0','" + dataNumber4 + "');";

                        cmd.ExecuteNonQuery();

                        primaryKey++;
                    }

                }//END WHILE
            }//END FOR

            #endregion

            connection.Close();

            connection.Dispose();
            GC.Collect();
            Thread.Sleep(5000);

        }
        public void SaveToPackingDBF(List<OrderModel> _checks, string _batchNumber, Main _mainForm)
        {
            string dbConnection;
            string tempCheckType = "";
            int blockNo = 0, blockCounter = 0;
            DbConServices db = new DbConServices();

            var listofchecks = _checks.Select(e => e.BankCode).Distinct().ToList();

            foreach (string checktype in listofchecks)
            {


                dbConnection = "Provider=VfpOleDB.1; Data Source=" + Application.StartupPath + "\\" + Main.outputFolder + "\\Packing.dbf" + "; Mode=ReadWrite;";

                //Check if packing file exists
                //if (!File.Exists(_filepath))
                //{
                string Sserial = "";
                string Eserial = "";
                OleDbConnection oConnect = new OleDbConnection(dbConnection);
                OleDbCommand oCommand;
                oConnect.Open();

                oCommand = new OleDbCommand("DELETE FROM PACKING", oConnect);
                oCommand.ExecuteNonQuery();
                foreach (var check in _checks)
                {
                    if (check.BankCode == null)
                    { }
                    else
                    {
                        Sserial = check.StartingSerial.ToString();
                        Eserial = check.EndingSerial.ToString();
                        while (Sserial.Length < 7)
                            Sserial = "0" + Sserial;

                        while (Eserial.Length < 7)
                            Eserial = "0" + Eserial;



                        if (tempCheckType != check.ChkType)
                            blockNo = 1;

                        tempCheckType = check.ChkType;

                        if (blockCounter < 4)
                            blockCounter++;
                        else
                        {
                            blockCounter = 1;
                            blockNo++;
                        }

                        string sql = "INSERT INTO PACKING (BATCHNO,BLOCK, RT_NO,BRANCH, ACCT_NO, ACCT_NO_P, CHKTYPE, ACCT_NAME1,ACCT_NAME2," +
                         "NO_BKS, CK_NO_B, CK_NO_E, DELIVERTO, CHKNAME) VALUES('" + _batchNumber + "'," + blockNo.ToString() + ",' ','" + check.BankName +
                         "',' ',' ','" + check.ChkType + "',' ',' ',1,'" +
                        Sserial + "','" + Eserial + "',' ',' ')";

                        oCommand = new OleDbCommand(sql, oConnect);

                        oCommand.ExecuteNonQuery();
                    }
                }
                oConnect.Close();

            }
        }

        public static void OfficeMDB(List<OrderModel> _orders, string _batch, string _chkType, string _fileName)
        {
            string txtName = "";
            string fileName = "";
            if (_chkType == "CS" || _chkType == "CV")
            {
                fileName = Application.StartupPath + "\\MDB\\" + _fileName;
                //  txtName = "CS" + _batch.Substring(0, 4) + "C.txt";
            }
            //else if (_chkType == "SR")
            //{
            //    fileName = Application.StartupPath + "\\" + Main.outputFolder + "\\" + _fileName;
            //    txtName = "ISL" + _batch.Substring(0, 4) + "S.20P";
            //}

            FileInfo fileInfo = new FileInfo(fileName);

            if (fileInfo.Exists)
                fileInfo.Delete();

            string mdbConn = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + fileName + "; Jet OLEDB:Engine Type=5";

            Catalog tClass = new Catalog();

            tClass.Create(mdbConn);

            System.Runtime.InteropServices.Marshal.ReleaseComObject(tClass);

            GC.Collect();

            tClass = new Catalog();

            tClass.let_ActiveConnection(mdbConn);

            //for (int x = 1; x <= 4; x++)
            //{

                Table tTable = new Table();

               // string tableTemp = "";

              //  if (x == 1)
                 //   tableTemp = "Out";
                //else
                //    tableTemp = "Outs";

                tTable.Name =  "Master_Database";


                //COLUMN NAMES
                //tTable.Columns.Append("BRSTN", DataTypeEnum.adVarWChar, 10);

                //tTable.Columns.Append("AccountNumber", DataTypeEnum.adVarWChar, 30);

                //tTable.Columns.Append("RT1to5", DataTypeEnum.adVarWChar, 5);

                //tTable.Columns.Append("RT6to9", DataTypeEnum.adVarWChar, 5);

                //tTable.Columns.Append("AccountNumberWithHypen", DataTypeEnum.adVarWChar, 30);

                //tTable.Columns.Append("Serial", DataTypeEnum.adVarWChar, 10);

                //tTable.Columns.Append("Name1", DataTypeEnum.adVarWChar, 50);

                //tTable.Columns.Append("Name2", DataTypeEnum.adVarWChar, 50);

                //tTable.Columns.Append("Name3", DataTypeEnum.adVarWChar, 50);

                //tTable.Columns.Append("Address1", DataTypeEnum.adVarWChar, 100);

                //tTable.Columns.Append("Address2", DataTypeEnum.adVarWChar, 100);

                //tTable.Columns.Append("Address3", DataTypeEnum.adVarWChar, 100);

                //tTable.Columns.Append("Address4", DataTypeEnum.adVarWChar, 100);

                //tTable.Columns.Append("Address5", DataTypeEnum.adVarWChar, 100);

                tTable.Columns.Append("Batch", DataTypeEnum.adVarWChar, 100);

                tTable.Columns.Append("BankName", DataTypeEnum.adVarWChar, 100);

                tTable.Columns.Append("BranchCode", DataTypeEnum.adVarWChar, 100);

                tTable.Columns.Append("Department", DataTypeEnum.adVarWChar, 100);

                tTable.Columns.Append("Serial", DataTypeEnum.adVarWChar, 10);

             //   tTable.Columns.Append("SerialNumber", DataTypeEnum.adVarWChar, 100);

                tTable.Columns.Append("StartingSerial", DataTypeEnum.adVarWChar, 20);

                tTable.Columns.Append("EndingSerial", DataTypeEnum.adVarWChar, 20);

                tTable.Columns.Append("Quantity", DataTypeEnum.adVarWChar, 30);

                tTable.Columns.Append("PcsPerBook", DataTypeEnum.adVarWChar, 3);

                tTable.Columns.Append("FileName", DataTypeEnum.adVarWChar, 30);

                tTable.Columns.Append("PrimaryKey", DataTypeEnum.adVarWChar);

                tTable.Columns["PrimaryKey"].Attributes = ColumnAttributesEnum.adColNullable;

                tTable.Columns.Append("PageNumber", DataTypeEnum.adVarWChar);

                tTable.Columns.Append("DataNumber", DataTypeEnum.adVarWChar, 20);

                tClass.Tables.Append((object)tTable);
           // }//END FOR

            GC.Collect(); //IF NOT INCLUDED FILE REMAIN IN OPENED STATUS

            OleDbConnection connection = new OleDbConnection(mdbConn);

            OleDbCommand cmd = new OleDbCommand();

            cmd.Connection = connection;

            connection.Open();

            int primaryKey = 1;
            int datacount = 0;


            // int dataNumber1 = 0, dataNumber2 = 0, dataNumber3 = 0, dataNumber4 = 0;//SERVE AS DATANUMBER
            #region 1Out Format
            //1OUTS FORMAT
            //foreach (var check in _orders)
            //{
            //  string RTFirst = check.BRSTN.Substring(0, 5);

            //  string RTLast = check.BRSTN.Substring(check.BRSTN.Length - 4, 4);
            foreach (var check in _orders)
            {
                datacount++;
            }

                string startSeries = _orders[0].StartingSerial.ToString();

                string endSeries = _orders[datacount -1].EndingSerial.ToString();

                //  string acctNoHypen = check.AccountNo.Substring(0, 3) + "-" + check.AccountNo.Substring(3, 6) + "-" + check.AccountNo.Substring(9, 3);

                while (startSeries.Length < 7)
                    startSeries = "0" + startSeries;

                while (endSeries.Length < 7)
                    endSeries = "0" + endSeries;

                //Int64 endS = check.EndingSerial;
                //Int64 start = check.StartingSerial;
   
                   // string temp = start.ToString();

                    //while (temp.Length < 10)
                    //    temp = "0" + temp;
            cmd.CommandText = "INSERT INTO Master_Database (Batch,BankName,BranchCode,Department, Serial,StartingSerial, EndingSerial, Quantity," +
             "PcsPerBook, FileName, PrimaryKey, PageNumber, DataNumber) " +
             "VALUES ('" + _batch + "','" + _orders[0].BankName.ToUpper() + "','" + _orders[0].BranchCode + "','" + _orders[0].Department + "','','" + startSeries + "','" + endSeries +
             "','"+_orders.Count+"','50','" + txtName + "','" + primaryKey + "','0','" + primaryKey + "');";


           // start++;

                    primaryKey++;
                //}//END WHILE
          //  }//END FOR
            #endregion
          
            cmd.ExecuteNonQuery();

            connection.Close();

            connection.Dispose();

            GC.Collect();
            Thread.Sleep(5000);

        }

    }
}

