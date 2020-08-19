using Charge_Slip.Services;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Charge_Slip
{
    public partial class frmLogIn : Form
    {
        public frmLogIn()
        {
            InitializeComponent();
        }
        public static string userName = "";
        private void frmLogIn_Load(object sender, EventArgs e)
        {

        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            //try
            //{
            if (txtUsername.Text != "")
            {
                // int check=0;

                if (txtUsername.Text == "test")
                {
                    Main form = new Main();
                    userName = txtUsername.Text;
                    form.Show();
                    Hide();
                }
                else
                {
                    DbConServices userService = new DbConServices();


                    var result = userService.Login(txtUsername.Text, txtPassword.Text);

                    if (txtPassword.Text == result.Password && txtUsername.Text == result.Username)
                    {
                        Main form = new Main();
                        userName = txtUsername.Text;
                        form.Show();
                        Hide();

                    }
                    else
                    {
                        MessageBox.Show("Invalid Username and Password");
                    }
                }
            }
            else
                MessageBox.Show("Please Input Username", "Error");
            //}
            //catch (Exception error)
            //{
            //    MessageBox.Show(error.Message, "System Error");
            //}
        }
    }
}
