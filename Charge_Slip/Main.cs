using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net.Sockets;
using System.Security.AccessControl;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Charge_Slip.Models;
using Charge_Slip.Services;
using MySql.Data.MySqlClient;

namespace Charge_Slip
{
    public partial class Main : Form
    {
        private List<OrderModel> orderList = new List<OrderModel>();
        DbConServices con = new DbConServices();
        BankModel bank = new BankModel(); 
        string newSeries = "";
        public string batchfile = "";
        public DateTime deliveryDate;
        DateTime dateTime;
        public static string outputFolder = "";
        public string FileName = "";
        private string PreFixName = "";
        public string checktype = "CS";
        public string packingOutputFolder = "";
        public string printerFileOutputFolder = "";
        public int PcsPerBook;

        public Main()
        {
            InitializeComponent();
            dateTime = dateTimePicker1.MinDate = DateTime.Now; //Disable selection of backdated dates to prevent errors  
        }
       
        private void Main_Load(object sender, EventArgs e)
        {
            LoadBanks();
        }
        private void LoadBanks()
        {
            cmbBank.Items.Add("Banko Nuestra");
            cmbBank.Items.Add("Maybank");
            cmbBank.Items.Add("Metro Bank");
            cmbBank.Items.Add("Security Bank");
            cmbBank.Items.Add("Union Bank");    
            cmbBank.SelectedIndex = 0;
        }

        private void cmbBank_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void txtQty_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtQty_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void addToolStripMenuItem_Click(object sender, EventArgs e)
        {


           
        }
        private void Clear()
        {
            txtQty.Text = "";
            txtLastNo.Text = "";
        }

        private void cmbBank_SelectedIndexChanged(object sender, EventArgs e)
        {

           
            if (cmbBank.Text == "Metro Bank")
            {
                bank.tableName = "mbtc";
               bank.IpAddress = "192.168.0.241";
                //  bank.IpAddress = "localhost";
                PcsPerBook = 49;
                checktype = "CS";
                con.GetLastno(bank, checktype);
                txtLastNo.Text = bank.LastNo.ToString();
                
                //txtBranchCode.Text = bank.BranchCode;
                //txtDept.Text = bank.Department;

            }
            else if (cmbBank.Text == "Maybank")
            {
                bank.tableName = "maybank";
                // bank.IpAddress = "localhost";
                bank.IpAddress = "192.168.0.254";
                checktype = "CS";
                PcsPerBook = 49;
                con.GetLastno(bank, checktype);
                txtLastNo.Text = bank.LastNo.ToString();
                //txtBranchCode.Text = bank.BranchCode;
                //txtDept.Text = bank.Department;

            }
            else if(cmbBank.Text == "Union Bank")
            {
                bank.tableName = "union";
                // bank.IpAddress = "localhost";
                bank.IpAddress = "192.168.0.254";
                checktype = "CV";
                PcsPerBook = 99;
                con.GetLastno(bank, checktype);
                txtLastNo.Text = bank.LastNo.ToString();
                //txtBranchCode.Text = bank.BranchCode;
                //txtDept.Text = bank.Department;

            }
            else if (cmbBank.Text == "Security Bank")
            {
                bank.tableName = "sbtc";
                // bank.IpAddress = "localhost";
                bank.IpAddress = "192.168.0.254";
               
                checktype = "CS";
                con.GetLastno(bank, checktype);
                txtLastNo.Text = bank.LastNo.ToString();

                //txtBranchCode.Text = bank.BranchCode;
               // txtDept.Text = bank.Department;

            }
            else if(cmbBank.Text == "Banko Nuestra")
            {
                bank.tableName = "banko_nuestra";
                //bank.IpAddress = "localhost";
                bank.IpAddress = "192.168.0.254";
           
                checktype = "CS";
                con.GetLastno(bank, checktype);
                txtLastNo.Text = bank.LastNo.ToString();
                //txtBranchCode.Text = bank.BranchCode;
                //txtDept.Text = bank.Department;
              
            }
        }
        public static DialogResult InputBox(string title, string promptText, ref string value)
        {

            Form form = new Form();
            Label label = new Label();
            TextBox textBox = new TextBox();
            Button buttonOk = new Button();
            Button buttonCancel = new Button();

            form.Text = title;
            label.Text = promptText;
            textBox.Text = value;

            buttonOk.Text = "OK";
            buttonCancel.Text = "Cancel";
            buttonOk.DialogResult = DialogResult.OK;
            buttonCancel.DialogResult = DialogResult.Cancel;

            label.SetBounds(9, 20, 372, 13);
            textBox.SetBounds(12, 36, 372, 20);
            buttonOk.SetBounds(228, 72, 75, 23);
            buttonCancel.SetBounds(309, 72, 75, 23);

            label.AutoSize = true;
            textBox.Anchor = textBox.Anchor | AnchorStyles.Right;
            buttonOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            buttonCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;

            form.ClientSize = new Size(396, 107);
            form.Controls.AddRange(new Control[] { label, textBox, buttonOk, buttonCancel });
            form.ClientSize = new Size(Math.Max(300, label.Right + 10), form.ClientSize.Height);
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.StartPosition = FormStartPosition.CenterScreen;
            form.MinimizeBox = false;
            form.MaximizeBox = false;
            form.AcceptButton = buttonOk;
            form.CancelButton = buttonCancel;

            DialogResult dialogResult = form.ShowDialog();
            value = textBox.Text;
            return dialogResult;
        }

        private void generateToolStripMenuItem1_Click(object sender, EventArgs e)
        {
          
        }

        private void txtLastNo_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void addToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            //     BankModel bank = new BankModel();
            Int64 SN;
            Int64 EN;
            if (txtQty.Text != "")
            {
                int qty = int.Parse(txtQty.Text);


                //if (txtLastNo.Text == "0")
                //{
                ///InputBox("Serial Number", "Please Input Serial Number ", ref newSeries);
                //    SN = Int64.Parse(newSeries) + 1;

                // }
                //else
                // {
                SN = Int64.Parse(txtLastNo.Text) + 1;
                //order.StartingSerial = SN;
                //EN =SN + 50;
                //  order.EndingSerial = EN;
                //  }
                for (int i = 0; i < qty; i++)
                {
                    OrderModel order = new OrderModel();
                    if (cmbBank.Text == "Metro Bank")
                    {
                        order.BankCode = "mbtc";
                        outputFolder = "Metro_bank";
                        packingOutputFolder = "METRO";
                        printerFileOutputFolder = "METRO";
                        PreFixName = "MB";
                        order.ChkType = "CS";
                        order.ChequeName = "Charge Slip";

                    }
                    else if (cmbBank.Text == "Union Bank")
                    {
                        order.BankCode = "union";
                        outputFolder = "Union_bank";
                        packingOutputFolder = "UNION";
                        printerFileOutputFolder = "UNION";
                        PreFixName = "UB";
                        order.ChkType = "CV";
                        checktype = "CV";
                        order.ChequeName = "Voucher";
                    }
                    else if (cmbBank.Text == "Security Bank")
                    {
                        order.BankCode = "sbtc";
                        outputFolder = "Security_bank";
                        packingOutputFolder = "SBTC";
                        printerFileOutputFolder = "SBTC";
                        PreFixName = "SB";
                        order.ChkType = "CS";
                        order.ChequeName = "Charge Slip";

                    }
                    else if (cmbBank.Text == "Maybank")
                    {
                        order.BankCode = "maybank";
                        outputFolder = "Maybank";
                        packingOutputFolder = "MAYBANK";
                        printerFileOutputFolder = "MAYBANK";
                        PreFixName = "M";
                        order.ChkType = "CS";
                        order.ChequeName = "Charge Slip";

                    }
                    else
                    {
                        order.BankCode = "banko_nuestra";
                        outputFolder = "Banko_nuestra";
                        packingOutputFolder = "Banko_Nuestra";
                        printerFileOutputFolder = "Banko_Nuestra";
                        PreFixName = "BN";
                        order.ChkType = "CS";
                        order.ChequeName = "Charge Slip";
                    }

                    order.Quantity = 1;
                    
                   
                    order.BankName = cmbBank.Text;
                    order.Department = txtDept.Text;
                    order.BranchCode = txtBranchCode.Text;


                    //order.StartingSerial = Int64.Parse(lblLastSeries.Text);
                    order.StartingSerial = SN.ToString();
                    EN = SN + PcsPerBook;
                    order.EndingSerial = EN.ToString();

                    orderList.Add(order);
                    SN = EN + 1;


                }


                BindingSource checkBind = new BindingSource();
                checkBind.DataSource = orderList;
                dataGridView1.DataSource = checkBind;
                Clear();
                generateToolStripMenuItem1.Enabled = true;
                cmbBank.Enabled = false;
                // lblTotal.Text = orderList.Count.ToString();
                //   MessageBox.Show("No Errors Found", "System Message");
            }
            else
                MessageBox.Show("Please enter how many booklets!!!");
        }

        private void generateToolStripMenuItem1_Click_1(object sender, EventArgs e)
        {
            ProcessServices proc = new ProcessServices();
            ZipServices zip = new ZipServices();
            deliveryDate = dateTimePicker1.Value;
            if (deliveryDate == dateTime)
            {
                MessageBox.Show("Please set Delivery Date!");
            }
            else
            {
                deliveryDate = dateTimePicker1.Value;
                if (txtBatch.Text != "")
                {
                    batchfile = txtBatch.Text;
                    FileName = PreFixName + batchfile.Substring(0, 4) + checktype;
                    proc.DoBlockProcess(orderList, this);
                    proc.SaveToPackingDBF(orderList, batchfile, this);
                    zip.DeleteFiles(".mdb",Application.StartupPath +"\\" +Main.outputFolder);
                    zip.DeleteFiles(".mdb", Application.StartupPath + "\\MDB");
                    //zip.DeleteMDB(".mdb");
                    ProcessServices.CreteaMDB(orderList, batchfile, checktype, FileName + ".mdb");

                    ProcessServices.OfficeMDB(orderList, batchfile, checktype, "ChargeSlip.mdb");
                    ZipServices.CopyMDB(frmLogIn.userName, this, FileName);
                    ZipServices.CopyMDB(frmLogIn.userName, this, "ChargeSlip");
                    zip.CopyPacking(frmLogIn.userName, this);
                    zip.CopyMDB(frmLogIn.userName, this);


                    Thread.Sleep(5000);
                    con.DumpMySQL(bank);
                    zip.ZipFileS(frmLogIn.userName, this);
                 
                    if (bank.IpAddress == "192.168.0.254")
                    {
                        zip.CopyZipFile(frmLogIn.userName, this, bank.IpAddress);
                    }
                    else
                        zip.CopyZipFileM(frmLogIn.userName, this, bank.IpAddress);

                    zip.DeleteFiles(".SQL", Application.StartupPath + "\\" + outputFolder);
                    for (int i = 0; i < orderList.Count; i++)
                    {
                        con.SavedDatatoDatabase(orderList[i], batchfile, bank.tableName, bank.IpAddress, deliveryDate);
                    }
                  
                    MessageBox.Show("Done!!");
                    GC.Collect();
                    Environment.Exit(0);
                }
                else
                    MessageBox.Show("Please Enter Batch Number!!!");
            }

        }

        private void txtDept_TextChanged(object sender, EventArgs e)
        {
            txtDept.CharacterCasing = CharacterCasing.Upper;
        }
    }
}
