using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BAtest
{
    public partial class Report : Form
    {
        public Report()
        {
            InitializeComponent();
        }

        OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=All Dept.accdb");
        
        private void button1_Click(object sender, EventArgs e)
        {
            con.Open();
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();
            dataGridView1.Refresh();
            OleDbCommand cmd = new OleDbCommand("select Date, Name, Designation, Department, GenerateTime from Report WHERE Date = cdate(date())", con);
            OleDbDataReader rd = cmd.ExecuteReader();
            DataTable dt = new DataTable();
            dt.Load(rd);
            dataGridView1.DataSource = dt;
            this.dataGridView1.DefaultCellStyle.Font = new Font("Tahoma", 12);
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.Columns[0].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            dataGridView1.Columns[1].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            dataGridView1.Columns[2].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            dataGridView1.Columns[3].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            dataGridView1.Columns[4].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            dataGridView1.Columns[5].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            DataGridViewColumn sn = dataGridView1.Columns[0];
            sn.Width = 10;
            DataGridViewColumn cd = dataGridView1.Columns[1];
            cd.Width = 25;
            DataGridViewColumn emp = dataGridView1.Columns[2];
            emp.Width = 55;
            DataGridViewColumn des = dataGridView1.Columns[3];
            des.Width = 25;
            DataGridViewColumn dept = dataGridView1.Columns[4];
            dept.Width = 30;
            DataGridViewColumn tm = dataGridView1.Columns[5];
            tm.Width = 110;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.White;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.Black;
            con.Close();
        }
        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            dataGridView1.Rows[e.RowIndex].Cells[0].Value = (e.RowIndex + 1).ToString();
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            showMonthlyReport(comboBox1.SelectedIndex + 1);
        }
        private void showMonthlyReport(int m)
        {
            con.Open();
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();
            dataGridView1.Refresh();
            OleDbCommand cmd = new OleDbCommand("select Date, Name, Designation, Department, GenerateTime from Report WHERE month(Date) = " + m + "", con);
            OleDbDataReader rd = cmd.ExecuteReader();
            DataTable dt = new DataTable();
            dt.Load(rd);
            dataGridView1.DataSource = dt;
            this.dataGridView1.DefaultCellStyle.Font = new Font("Tahoma", 12);
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.Columns[0].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            dataGridView1.Columns[1].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            dataGridView1.Columns[2].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            dataGridView1.Columns[3].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            dataGridView1.Columns[4].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            dataGridView1.Columns[5].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            DataGridViewColumn sn = dataGridView1.Columns[0];
            sn.Width = 10;
            DataGridViewColumn cd = dataGridView1.Columns[1];
            cd.Width = 25;
            DataGridViewColumn emp = dataGridView1.Columns[2];
            emp.Width = 55;
            DataGridViewColumn des = dataGridView1.Columns[3];
            des.Width = 25;
            DataGridViewColumn dept = dataGridView1.Columns[4];
            dept.Width = 30;
            DataGridViewColumn tm = dataGridView1.Columns[5];
            tm.Width = 110;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.White;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.Black;
            con.Close();
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selecteddept = (string)comboBox2.SelectedItem;
            con.Open();
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();
            dataGridView1.Refresh();
            OleDbCommand cmd;
            if (selecteddept.Equals("AllDept"))
            {
                cmd = new OleDbCommand("select Date, Name, Designation, Department, GenerateTime  from Report WHERE Date BETWEEN #" + dateTimePicker1.Value.Date + "# AND #" + dateTimePicker2.Value.Date + "#", con);
            }
            else
            {
                cmd = new OleDbCommand("select Date, Name, Designation, Department, GenerateTime from Report WHERE Department = '" + selecteddept + "' and Date BETWEEN #" + dateTimePicker1.Value.Date + "# AND #" + dateTimePicker2.Value.Date + "#" , con);                
            }
            OleDbDataReader rd = cmd.ExecuteReader();
            DataTable dt = new DataTable();
            dt.Load(rd);
            dataGridView1.DataSource = dt;
            this.dataGridView1.DefaultCellStyle.Font = new Font("Tahoma", 12);
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.Columns[0].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            dataGridView1.Columns[1].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            dataGridView1.Columns[2].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            dataGridView1.Columns[3].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            dataGridView1.Columns[4].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            dataGridView1.Columns[5].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            DataGridViewColumn sn = dataGridView1.Columns[0];
            sn.Width = 10;
            DataGridViewColumn cd = dataGridView1.Columns[1];
            cd.Width = 25;
            DataGridViewColumn emp = dataGridView1.Columns[2];
            emp.Width = 55;
            DataGridViewColumn des = dataGridView1.Columns[3];
            des.Width = 25;
            DataGridViewColumn dept = dataGridView1.Columns[4];
            dept.Width = 30;
            DataGridViewColumn tm = dataGridView1.Columns[5];
            tm.Width = 110;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.White;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.Black;
            con.Close();
        }
    }
}
