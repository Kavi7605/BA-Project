using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BAtest
{
    public partial class Login : Form
    {
        private MainForm frm;
        public Login(MainForm frm)
        {
            InitializeComponent();
            this.frm = frm;
        }
        string username = "Admin";
        string pass = "adminpassword";
        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == username && textBox2.Text == pass)
            {
                regenerate rg = new regenerate();
                rg.Show();
            }
            else 
            {
                string title = "Wrong User Credentials";
                string msg = "Wrong Username or Password. \nRe-enter Correct Credentials.";
                MessageBox.Show(msg, title);
            }
            textBox1.Text = "";
            textBox2.Text = "";
        }
        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                if (textBox1.Text == username && textBox2.Text == pass)
                {
                    frm.enableAdminMode();
                    string title = "Admin Mode";
                    string msg = "You are now in Admin Mode. \nYou can Access Update and Delete Button from Edit window.";
                    MessageBox.Show(title, msg);
                }
                else
                {
                    string title = "Wrong User Credentials";
                    string msg = "Wrong Username or Password. \nRe-enter Correct Credentials.";
                    MessageBox.Show(msg, title);
                }
            }
            catch (Exception)
            {
                MessageBox.Show("You cant enable delete button now. Open the Edit window to enable it.");
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (frm.AdminMode)
                {
                    string message = "Are you sure you want to Logout?";
                    string title = "Logout window";
                    MessageBoxButtons buttons = MessageBoxButtons.YesNo;
                    DialogResult result = MessageBox.Show(message, title, buttons);
                    if (result == DialogResult.Yes)
                    {
                        frm.logoutAdmin();
                    }
                }
                else
                {
                    string message = "You are already Logged out.";
                    string title = "Logout window";
                    MessageBox.Show(message, title);
                }
            }
            catch(Exception e1)
            {
                MessageBox.Show(e1.Message);
            }
        }

    }
}
