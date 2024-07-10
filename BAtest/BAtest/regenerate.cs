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
    public partial class regenerate : Form
    {
        public regenerate()
        {
            InitializeComponent();
        }
        OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=All Dept.accdb");
                
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                textBox1.Text = dataGridView1.CurrentRow.Cells[0].Value.ToString();
                String deleteData = "DELETE FROM Report WHERE ID = ?";
                con.Open();
                OleDbCommand cmd;
                cmd = new OleDbCommand(deleteData, con);
                cmd.Parameters.AddWithValue("ID", OleDbType.VarChar).Value = Convert.ToInt16(textBox1.Text);
                cmd.ExecuteNonQuery();
                con.Close();
                data();
            }
            catch (Exception e1)
            {
                MessageBox.Show(e1.Message);
            }
        }

        private void data()
        {
            con.Open();
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();
            dataGridView1.Refresh();
            OleDbCommand cmd = new OleDbCommand("select ID, Department from Report WHERE Date = cdate(date())", con);
            OleDbDataReader rd = cmd.ExecuteReader();
            DataTable dt = new DataTable();
            dt.Load(rd);
            dataGridView1.DataSource = dt;
            this.dataGridView1.DefaultCellStyle.Font = new Font("Tahoma", 12);
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.Columns[0].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            dataGridView1.Columns[1].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            DataGridViewColumn id = dataGridView1.Columns[0];
            DataGridViewColumn dept = dataGridView1.Columns[1];
            id.Width = 50;
            dept.Width = 272;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.White;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.Black;
            con.Close();
        }

        private void regenerate_Load(object sender, EventArgs e)
        {
            data();
        }
        private void dataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            textBox1.Text = dataGridView1.CurrentRow.Cells[0].Value.ToString();
        }
    }
}
