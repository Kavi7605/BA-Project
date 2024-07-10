using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BAtest
{
    public partial class printpdf : Form
    {
        public printpdf()
        {
            InitializeComponent();
        }
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == (Keys.Control | Keys.P))
            {
                Print(this.panel1);
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
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

        OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=All Dept.accdb");
        private void printpdf_Load(object sender, EventArgs e)
        {
            con.Open();
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();
            dataGridView1.Refresh();
            OleDbCommand cmd = new OleDbCommand("SELECT Name, EmpNo, Department FROM Report WHERE Date = cdate(date())", con);
            OleDbDataReader rd = cmd.ExecuteReader();
            DataTable dt = new DataTable();
            dt.Load(rd);
            dataGridView1.DataSource = dt;
            dataGridView1.Columns.Add("newColumnName", "Column Name in Text");
            dataGridView1.Columns.Add("newColumnName1", "Column Name in Text");
            dataGridView1.Columns.Add("newColumnName2", "Column Name in Text");
            dataGridView1.Columns.Add("newColumnName3", "Column Name in Text");
            dataGridView1.Columns.Add("newColumnName4", "Column Name in Text");
            dataGridView1.Columns.Add("newColumnName5", "Column Name in Text");
            dataGridView1.Columns.Add("newColumnName6", "Column Name in Text");
            this.dataGridView1.DefaultCellStyle.Font = new Font("Tahoma", 12);
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.Columns[0].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            dataGridView1.Columns[1].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            dataGridView1.Columns[2].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            dataGridView1.Columns[3].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            dataGridView1.Columns[4].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            dataGridView1.Columns[5].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            dataGridView1.Columns[6].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            dataGridView1.Columns[7].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            dataGridView1.Columns[8].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            dataGridView1.Columns[9].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            dataGridView1.Columns[10].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            DataGridViewColumn sn = dataGridView1.Columns[0];
            sn.Width = 47;
            DataGridViewColumn name1 = dataGridView1.Columns[1];
            name1.Width = 214;
            DataGridViewColumn emp = dataGridView1.Columns[2];
            emp.Width = 121;
            DataGridViewColumn dept = dataGridView1.Columns[3];
            dept.Width = 135;
            DataGridViewColumn col1 = dataGridView1.Columns[4];
            col1.Width = 105;
            DataGridViewColumn col2 = dataGridView1.Columns[5];
            col2.Width = 120;
            DataGridViewColumn col3 = dataGridView1.Columns[6];
            col3.Width = 106;
            DataGridViewColumn col4 = dataGridView1.Columns[7];
            col4.Width = 91;
            DataGridViewColumn col5 = dataGridView1.Columns[8];
            col5.Width = 119;
            DataGridViewColumn col6 = dataGridView1.Columns[9];
            col6.Width = 106;
            DataGridViewColumn col7 = dataGridView1.Columns[10];
            col7.Width = 91;
            dataGridView1.ColumnHeadersVisible = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.White;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.Black;
            con.Close();
        }

        private void dataGridView1_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            dataGridView1.Rows[e.RowIndex].Cells[0].Value = (e.RowIndex + 1).ToString();
            dataGridView1.Columns[0].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[1].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[2].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[3].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[4].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[5].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[6].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[7].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[8].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[9].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[10].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
        }
    }
}
