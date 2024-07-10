using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using Microsoft.Win32;

namespace BAtest{
    public partial class MainForm : Form {
        public MainForm()
        {
            InitializeComponent();
        }
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == (Keys.Control | Keys.H))
            {
                Login lg = new Login(this);
                lg.Show();
                return true;
            }
            if (keyData == (Keys.Control | Keys.D))
            {
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    row.Cells["PA"].Value = false;
                }
                return true;
            }
            if (keyData == (Keys.Control | Keys.S))
            {
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    row.Cells["PA"].Value = true;
                }
                return true;
            }
            if (keyData == (Keys.Control | Keys.P))
            {
                Print(this.panel1);
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }
        public void enableAdminMode()
        {
            AdminMode = true;
        }
        public void logoutAdmin()
        {
            AdminMode = false;
        }
        public Boolean AdminMode;

        string[] fivenames = new string[5];
        int i = -1;
        Boolean flag;
        Boolean flag1  = true;
        string[] deptnames = { "AIASL", "ATC", "BPCL", "CISF", "CNSIT", "FIREGUJSAIL", "MT" };
        Label[] labelArray;

        OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=All Dept.accdb");

        private void Style1()
        {
            this.dataGridView1.DefaultCellStyle.Font = new Font("Tahoma", 12);
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.Columns[0].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            dataGridView1.Columns[1].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            dataGridView1.Columns[2].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            dataGridView1.Columns[3].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            dataGridView1.Columns[4].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            DataGridViewColumn SN = dataGridView1.Columns[0];
            SN.Width = 25;
            DataGridViewColumn PA = dataGridView1.Columns[1];
            PA.Width = 30;
            DataGridViewColumn Name = dataGridView1.Columns[2];
            Name.Width = 200;
            DataGridViewColumn Designation = dataGridView1.Columns[3];
            Designation.Width = 170;
            DataGridViewColumn Mobile = dataGridView1.Columns[4];
            Mobile.Width = 140;
        }
        private void Style2()
        {
            this.dataGridView1.DefaultCellStyle.Font = new Font("Tahoma", 12);
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.Columns[0].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            dataGridView1.Columns[1].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            dataGridView1.Columns[2].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            dataGridView1.Columns[3].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            dataGridView1.Columns[4].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            dataGridView1.Columns[5].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            DataGridViewColumn SN = dataGridView1.Columns[0];
            SN.Width = 35;
            DataGridViewColumn PA = dataGridView1.Columns[1];
            PA.Width = 35;
            DataGridViewColumn EmpNo = dataGridView1.Columns[2];
            EmpNo.Width = 110;
            DataGridViewColumn Name = dataGridView1.Columns[3];
            Name.Width = 200;
            DataGridViewColumn Designation = dataGridView1.Columns[4];
            Designation.Width = 160;
            DataGridViewColumn Mobile = dataGridView1.Columns[5];
            Mobile.Width = 140;
        }
        private void oledb()
        {
            label4.Text = "";
            label8.Text = "";
            label9.Text = "";
            label10.Text = "";
            label12.Text = "";
            label13.Text = "";
            label14.Text = "";
            label15.Text = "";
            label16.Text = "";
            label17.Text = "";
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();
            dataGridView1.Refresh();
            con.Open();
            OleDbCommand cmd;
            if (flag == true)
            {
                cmd = new OleDbCommand("SELECT PA, Name, Designation, Mobile from AllDept WHERE Department = '" + deptnames[i] + "'", con);
            }
            else
            {
                cmd = new OleDbCommand("select PA, EmpNo, Name, Designation, Mobile from AllDept WHERE Department = '" + deptnames[i] + "'", con);
            }
            OleDbDataReader rd = cmd.ExecuteReader();
            DataTable dt = new DataTable();
            dt.Load(rd);
            dataGridView1.DataSource = dt;
            con.Close();
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.Black;
            UpdateButtonState();
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                string selecteddept = (string)comboBox1.SelectedItem;
                if (selecteddept == "AIASL")
                {
                    i = 0;
                    flag = true;
                    oledb();
                    dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.PeachPuff;
                    Style1();
                    label5.Text = "AIASL (AI Airport Service Limited)";
                    label5.ForeColor = Color.Black;
                    label5.BackColor = Color.PeachPuff;
                }
                if (selecteddept == "ATC")
                {
                    i = 1;
                    flag = false;
                    oledb();
                    dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Aqua;
                    Style2();
                    label5.Text = "ATC (Air Traffic Control)";
                    label5.ForeColor = Color.Black;
                    label5.BackColor = Color.Aqua;
                }
                if (selecteddept == "BPCL")
                {
                    i = 2;
                    flag = true;
                    oledb();
                    dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Yellow;
                    Style1();
                    label5.Text = "BPCL (Bharat Petroleum Corporation Limited)";
                    label5.ForeColor = Color.Black;
                    label5.BackColor = Color.Yellow;
                }
                if (selecteddept == "CISF")
                {
                    i = 3;
                    flag = true;
                    oledb();
                    dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.SaddleBrown;
                    Style1();
                    label5.Text = "CISF (Central Industrial Security Force) Driver";
                    label5.ForeColor = Color.White;
                    label5.BackColor = Color.SaddleBrown;
                }
                if (selecteddept == "CNSIT")
                {
                    i = 4;
                    flag = false;
                    oledb();
                    dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.LimeGreen;
                    Style2();
                    label5.Text = "CNS (Communication Navigation and Suveillance) & IT (Information Technology)";
                    label5.ForeColor = Color.Black;
                    label5.BackColor = Color.LimeGreen;
                }
                if (selecteddept == "FIREGUJSAIL")
                {
                    i = 5;
                    flag = true;
                    oledb();
                    dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Red;
                    Style1();
                    label5.Text = "FIRE (Air Fire Station(AFS)) & GUJSAIL(Gujarat State Aviation Infrastucture Company Ltd.)";
                    label5.ForeColor = Color.White;
                    label5.BackColor = Color.Red;
                }
                if (selecteddept == "MT")
                {
                    i = 6;
                    flag = true;
                    oledb();
                    dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.DarkOrchid;
                    Style1();
                    label5.Text = "MT (Motor Transport)";
                    label5.ForeColor = Color.Black;
                    label5.BackColor = Color.DarkOrchid;
                }
            }
            catch (Exception e1)
            {
                MessageBox.Show(e1.Message);
            }
        }
        private void Generate(object sender, EventArgs e) {
            try
            {
                if (flag1)
                {
                    string[] array = new string[dataGridView1.Rows.Count];
                    int j = 0;
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        bool isSelected = Convert.ToBoolean(row.Cells["PA"].Value);
                        if (isSelected)
                        {
                            array[j] = Convert.ToString(row.Cells["Name"].Value);
                            j++;
                        }
                    }
                    Random rand = new Random();
                    con.Open();
                    OleDbCommand cmd;
                    int k = 0;                    
                    double f = Convert.ToDouble(array.Length) / 10;
                    double ceiling = Math.Ceiling(f);
                    while (0 < f)
                    {
                        labelArray[k].Text = array[rand.Next(0, j)];
                        fivenames[k] = labelArray[k].Text;
                        cmd = new OleDbCommand("INSERT INTO Report SELECT Name, Department, EmpNo, Designation FROM AllDept WHERE Name = '" + labelArray[k].Text + "'", con);
                        cmd.ExecuteNonQuery();
                        cmd = new OleDbCommand("UPDATE Report SET GenerateTime='" + label3.Text + "' WHERE Name = '" + labelArray[k].Text + "'", con);
                        cmd.ExecuteNonQuery();
                        k++;
                        f--;
                    }
                    cmd = new OleDbCommand("SELECT Designation FROM Report WHERE Date = cdate(date()) AND Department = '" + deptnames[i] + "'", con);
                    OleDbDataReader rd = cmd.ExecuteReader();

                    if (rd.HasRows)
                    {
                        button8.Enabled = false;
                        while (rd.Read())
                        {
                            labelArray[k].Text = (string)rd.GetValue(0);
                            k++;
                        }
                    }
                    else
                    {
                        button8.Enabled = true;
                    }
                    button8.Enabled = false;
                    con.Close();
                }
                else
                {
                    string[] array = new string[dataGridView1.Rows.Count];
                    int j = 0;
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        bool isSelected = Convert.ToBoolean(row.Cells["PA"].Value);
                        if (isSelected)
                        {
                            array[j] = Convert.ToString(row.Cells["Name"].Value);
                            j++;
                        }
                    }
                    Random rand = new Random();

                    con.Open();
                    OleDbCommand cmd;
                    label4.Text = array[rand.Next(0, j)];
                    cmd = new OleDbCommand("INSERT INTO AllDeptgenerate SELECT EmpNo, Name, Designation, Mobile, Department FROM AllDept WHERE Name = '" + label4.Text +"'", con);
                    cmd.ExecuteNonQuery();
                    button8.Enabled = false;
                    con.Close();
                    con.Open();
                    cmd = new OleDbCommand("SELECT Designation FROM AllDeptgenerate WHERE Date = cdate(date()) AND Name = '" + label4.Text + "'", con);
                    OleDbDataReader rd = cmd.ExecuteReader();
                    rd.Read();
                    label8.Text = (string)rd.GetValue(0);
                    con.Close();
                }
            }
            catch (Exception e1) {
                MessageBox.Show(e1.Message);
            }
        }
        private void UpdateButtonState()
        {
            try
            {
                con.Open();
                OleDbCommand cmd = new OleDbCommand("SELECT Name FROM Report WHERE Date = cdate(date()) AND Department = '" + deptnames[i] + "'", con);
                OleDbDataReader rd = cmd.ExecuteReader();

                int k = 0;
                if (rd.HasRows)
                {
                    button8.Enabled = false;
                    while (rd.Read())
                    {
                        labelArray[k].Text = (string)rd.GetValue(0);
                        k++;
                        k++;
                    }
                }
                else
                {
                    button8.Enabled = true;
                }
                con.Close();
                con.Open();
                cmd = new OleDbCommand("SELECT Designation FROM Report WHERE Date = cdate(date()) AND Department = '" + deptnames[i] + "'", con);
                rd = cmd.ExecuteReader();

                k = 0;
                if (rd.HasRows)
                {
                    button8.Enabled = false;
                    while (rd.Read())
                    {
                        k++;
                        labelArray[k].Text = (string)rd.GetValue(0);
                        k++;
                    }
                }
                else
                {
                    button8.Enabled = true;
                }
                con.Close();
            }
            catch (Exception e1)
            {
                MessageBox.Show(e1.Message);
            }
        }
        private void button12_Click(object sender, EventArgs e)
        {
            Report rp = new Report();
            rp.Show();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            label4.Text = "";
            label8.Text = "";
            label9.Text = "";
            label10.Text = "";
            label12.Text = "";
            label13.Text = "";
            label14.Text = "";
            label15.Text = "";
            label16.Text = "";
            label17.Text = "";
            flag1 = false;
            label5.Text = "All Department";
            label5.ForeColor = Color.White;
            label5.BackColor = Color.Black;
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();
            dataGridView1.Refresh();
            con.Open();
            OleDbCommand cmd = new OleDbCommand("SELECT PA, Department, EmpNo, Name, Designation, Mobile FROM AllDept", con);
            OleDbDataReader rd = cmd.ExecuteReader();
            DataTable dt = new DataTable();
            dt.Load(rd);
            dataGridView1.DataSource = dt;
            con.Close();
            this.dataGridView1.DefaultCellStyle.Font = new Font("Tahoma", 12);
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.Columns[0].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            dataGridView1.Columns[1].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            dataGridView1.Columns[2].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            dataGridView1.Columns[3].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            dataGridView1.Columns[4].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            dataGridView1.Columns[5].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            dataGridView1.Columns[6].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            DataGridViewColumn sn = dataGridView1.Columns[0];
            sn.Width = 40;
            DataGridViewColumn pa = dataGridView1.Columns[1];
            pa.Width = 40;
            DataGridViewColumn dept = dataGridView1.Columns[2];
            dept.Width = 100;
            DataGridViewColumn emp = dataGridView1.Columns[3];
            emp.Width = 80;
            DataGridViewColumn name = dataGridView1.Columns[4];
            name.Width = 210;
            DataGridViewColumn des = dataGridView1.Columns[5];
            des.Width = 130;
            DataGridViewColumn mob = dataGridView1.Columns[6];
            mob.Width = 120;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

            con.Open();
            cmd = new OleDbCommand("SELECT Name FROM AllDeptgenerate WHERE Date = cdate(date())", con);
            rd = cmd.ExecuteReader();
            if (rd.HasRows)
            {
                rd.Read();
                label4.Text = (string)rd.GetValue(0);
                button8.Enabled = false;
                //TimeSpan remainingTime = (DateTime.Today.AddDays(1)) - (DateTime.Now);
                //label10.Text = $"Button will be enabled in {remainingTime.Hours} hours and {remainingTime.Minutes} minutes and {remainingTime.Seconds} seconds.";
            }
            else
            {
                button8.Enabled = true;
            }
            con.Close();
            con.Open();
            cmd = new OleDbCommand("SELECT Designation FROM AllDeptgenerate WHERE Date = cdate(date())", con);
            rd = cmd.ExecuteReader();
            if (rd.HasRows)
            {
                rd.Read();
                label8.Text = (string)rd.GetValue(0);                
                //TimeSpan remainingTime = (DateTime.Today.AddDays(1)) - (DateTime.Now);
                //label10.Text = $"Button will be enabled in {remainingTime.Hours} hours and {remainingTime.Minutes} minutes and {remainingTime.Seconds} seconds.";
            }
            con.Close();
        }
        private void Print_btn(object sender, EventArgs e){
            printpdf pdf = new printpdf();
            pdf.Show();
        }
        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            dataGridView1.Rows[e.RowIndex].Cells[0].Value = (e.RowIndex + 1).ToString();
        }
        private void Edit_btn(object sender, EventArgs e){
            DataEntry de = new DataEntry(this);
            de.Show();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            this.Size = Screen.PrimaryScreen.WorkingArea.Size;
            timer1.Start();
            label2.Text = DateTime.Now.ToLongDateString();
            label3.Text = DateTime.Now.ToLongTimeString();
            labelArray = new Label[] { label4, label8, label9, label12, label13, label10, label14, label15, label16, label17 };
            label8.Text = "Designation";
            label9.Text = "";
            label12.Text = "";
            label13.Text = "";
            label10.Text = "";
            label14.Text = "";
            label15.Text = "";
            label16.Text = "";
            label17.Text = "";
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            label2.Text = DateTime.Now.ToLongDateString();
            label3.Text = DateTime.Now.ToLongTimeString();
        }
        Bitmap memorying;
        private void Print(Panel pnl)
        {
            PrinterSettings ps = new PrinterSettings();
            panel1 = pnl;
            getprintarea(pnl);
            printDocument1.PrintPage += new PrintPageEventHandler(printDocument1_PrintPage);
            if (printDialog1.ShowDialog() == DialogResult.OK)
            {
                printDocument1.Print();
            }
        }
        private void printDocument1_PrintPage(object sender, PrintPageEventArgs e)
        {
            Rectangle pagearea = e.PageBounds;
            e.Graphics.DrawImage(memorying, 0, 0);
        }
        private void getprintarea(Panel pnl)
        {
            memorying = new Bitmap(pnl.Width, pnl.Height);
            pnl.DrawToBitmap(memorying, new Rectangle(0, 0, pnl.Width, pnl.Height));
        }
    }
}
