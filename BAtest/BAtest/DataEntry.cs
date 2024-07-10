using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace BAtest
{
    public partial class DataEntry : Form
    {
        private MainForm frm;
        public DataEntry(MainForm frm)
        {
            this.frm = frm;
            InitializeComponent();
        }

        OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=All Dept.accdb");

        int num;
        string[] deptnames = { "AIASL", "ATC", "BPCL", "CISF", "CNSIT", "FIREGUJSAIL", "MT" };

        private void Oledb()
        {
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();
            dataGridView1.Refresh();
            con.Open();
            OleDbCommand cmd = new OleDbCommand("select ID, EmpNo, Name, Designation, Mobile from AllDept WHERE Department ='" + deptnames[num] + "'", con);
            OleDbDataReader rd = cmd.ExecuteReader();
            DataTable dt = new DataTable();
            dt.Load(rd);
            con.Close();
            dataGridView1.DataSource = dt;
            this.dataGridView1.DefaultCellStyle.Font = new Font("Tahoma", 10);
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.Columns[0].HeaderCell.Style.Font = new Font("Tahoma", 11, FontStyle.Bold);
            dataGridView1.Columns[1].HeaderCell.Style.Font = new Font("Tahoma", 11, FontStyle.Bold);
            dataGridView1.Columns[2].HeaderCell.Style.Font = new Font("Tahoma", 11, FontStyle.Bold);
            dataGridView1.Columns[3].HeaderCell.Style.Font = new Font("Tahoma", 11, FontStyle.Bold);
            dataGridView1.Columns[4].HeaderCell.Style.Font = new Font("Tahoma", 11, FontStyle.Bold);
            dataGridView1.Columns[5].HeaderCell.Style.Font = new Font("Tahoma", 11, FontStyle.Bold);
            DataGridViewColumn SN = dataGridView1.Columns[0];
            DataGridViewColumn ID = dataGridView1.Columns[1];
            DataGridViewColumn EmpNo = dataGridView1.Columns[2];
            DataGridViewColumn Name1 = dataGridView1.Columns[3];
            DataGridViewColumn Designation = dataGridView1.Columns[4];
            DataGridViewColumn Mobile = dataGridView1.Columns[5];
            SN.Width = 40;
            ID.Width = 40;
            EmpNo.Width = 80;
            Name1.Width = 230;
            Designation.Width = 180;
            Mobile.Width = 140;
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                //ComboBox comboBox = (ComboBox)sender;
                string selecteddept = (string)comboBox1.SelectedItem;
                if (selecteddept == "AIASL")
                {
                    num = 0;
                    AIASLdataGrid();
                }
                if (selecteddept == "ATC")
                {
                    num = 1;
                    ATCdataGrid();
                }
                if (selecteddept == "BPCL")
                {
                    num = 2;
                    BPCLdataGrid();
                }
                if (selecteddept == "CISF")
                {
                    num = 3;
                    CISFdataGrid();
                }
                if (selecteddept == "CNSIT")
                {
                    num = 4;
                    CNSITdataGrid();
                }
                if (selecteddept == "FIREGUJSAIL")
                {
                    num = 5;
                    FIREGUJSAILdataGrid();
                }
                if (selecteddept == "MT")
                {
                    num = 6;
                    MTdataGrid();
                }
            }
            catch (Exception e1)
            {
                MessageBox.Show(e1.Message);
            }
        }
        private void AIASLdataGrid()
        {
            Oledb();
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.PeachPuff;
            label7.Text = "AIASL (AI Airport Service Limited)";
            label7.ForeColor = Color.Black;
            label7.BackColor = Color.PeachPuff;
        }
        private void ATCdataGrid()
        {
            Oledb();
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Aqua;
            label7.Text = "ATC (Air Traffic Control)";
            label7.ForeColor = Color.Black;
            label7.BackColor = Color.Aqua;
        }
        private void BPCLdataGrid()
        {
            Oledb();
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Yellow;
            label7.Text = "BPCL (Bharat Petroleum Corporation Limited)";
            label7.ForeColor = Color.Black;
            label7.BackColor = Color.Yellow;
        }
        private void CISFdataGrid()
        {
            Oledb();
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.SaddleBrown;
            label7.Text = "CISF (Central Industrial Security Force) Driver";
            label7.ForeColor = Color.White;
            label7.BackColor = Color.SaddleBrown;
        }
        private void CNSITdataGrid()
        {
            Oledb();
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.LimeGreen;
            label7.Text = "CNS (Communication Navigation and Suveillance) & IT (Information Technology)";
            label7.ForeColor = Color.Black;
            label7.BackColor = Color.LimeGreen;
        }
        private void FIREGUJSAILdataGrid()
        {
            Oledb();
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Red;
            label7.Text = "FIRE (Air Fire Station(AFS)) & GUJSAIL(Gujarat State Aviation Infrastucture Company Ltd.)";
            label7.ForeColor = Color.White;
            label7.BackColor = Color.Red;
        }
        private void MTdataGrid()
        {
            Oledb();
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.DarkOrchid;
            label7.Text = "MT (Motor Transport)";
            label7.ForeColor = Color.Black;
            label7.BackColor = Color.DarkOrchid;
        }
        private void Deletedb()
        {
            string message = "Are you sure you want to Delete this data?";
            string title = "Delete Window";
            MessageBoxButtons buttons = MessageBoxButtons.YesNo;
            DialogResult result = MessageBox.Show(message, title, buttons);
            if (result == DialogResult.Yes)
            {
                String deleteData = "DELETE FROM AllDept WHERE ID = ?";
                con.Open();
                OleDbCommand cmd;
                cmd = new OleDbCommand(deleteData, con);
                cmd.Parameters.AddWithValue("ID", OleDbType.VarChar).Value = Convert.ToInt16(ID.Text);
                cmd.ExecuteNonQuery();
                con.Close();
            }
        }
        private void Deletebtn_Click(object sender, EventArgs e)
        {
            try
            {
                if (num == 0)
                {
                    Deletedb();
                    AIASLdataGrid();
                }
                else if (num == 1)
                {
                    Deletedb();
                    ATCdataGrid();
                }
                else if (num == 2)
                {
                    Deletedb();
                    BPCLdataGrid();
                }
                else if (num == 3)
                {
                    Deletedb();
                    CISFdataGrid();
                }
                else if (num == 4)
                {
                    Deletedb();
                    CNSITdataGrid();
                }
                else if (num == 5)
                {
                    Deletedb();
                    FIREGUJSAILdataGrid();
                }
                else if (num == 6)
                {
                    Deletedb();
                    MTdataGrid();
                }
                else
                {
                    MessageBox.Show("Else Condition is Excecuted.");
                }

            }
            catch (Exception e1)
            {
                con.Close();
                MessageBox.Show(e1.Message);
            }
        }
        private void clearInputFields()
        {
            try
            {
                NametextBox.Text = string.Empty;
                Designation.Text = string.Empty;
                Mobile.Text = string.Empty;
                ID.Text = string.Empty;
                EmpNo.Text = string.Empty;
            }
            catch (Exception e1)
            {
                MessageBox.Show(e1.Message);
            }
        }
        private void Updatedb()
        {
            String insertdata = "UPDATE AllDept SET EmpNo=?, Name=?, Designation=?, Mobile=? WHERE ID=?";
            con.Open();
            OleDbCommand cmd;
            cmd = new OleDbCommand(insertdata, con);
            if (string.IsNullOrWhiteSpace(EmpNo.Text))
            {
                cmd.Parameters.Add(new OleDbParameter { Value = DBNull.Value });
            }
            else
            {
                cmd.Parameters.Add(new OleDbParameter { Value = EmpNo.Text });
            }
            cmd.Parameters.Add(new OleDbParameter { Value = NametextBox.Text });
            cmd.Parameters.Add(new OleDbParameter { Value = Designation.Text });
            cmd.Parameters.Add(new OleDbParameter { Value = Mobile.Text });
            cmd.Parameters.Add(new OleDbParameter { Value = ID.Text });
            int dataInserted = cmd.ExecuteNonQuery();
            if(dataInserted > 0)
            {
                MessageBox.Show("Data inserted into Database.");
            }
            con.Close();
        }
        private void Updatebtn_Click(object sender, EventArgs e)
        {
            try
            {
                if (num == 0)
                {
                    Updatedb();
                    AIASLdataGrid();
                }
                else if (num == 1)
                {
                    Updatedb();
                    ATCdataGrid();
                }
                else if (num == 2)
                {
                    Updatedb();
                    BPCLdataGrid();
                }
                else if (num == 3)
                {
                    Updatedb();
                    CISFdataGrid();
                }
                else if (num == 4)
                {
                    Updatedb();
                    CNSITdataGrid();
                }
                else if (num == 5)
                {
                    Updatedb();
                    FIREGUJSAILdataGrid();
                }
                else if (num == 6)
                {
                    Updatedb();
                    MTdataGrid();
                }
                else
                {
                    MessageBox.Show("Else Condition is Excecuted.");
                }
            }
            catch (Exception e1)
            {
                con.Close();
                MessageBox.Show(e1.Message);
            }
        }
        private void Insertdb()
        {
            try
            {
                con.Open();
                OleDbCommand cmd;
                cmd = new OleDbCommand("INSERT INTO AllDept(EmpNo, Name, Designation, Mobile) VALUES(@EmpNo, @Name, @Designation, @Mobile)", con);

                if (string.IsNullOrWhiteSpace(EmpNo.Text))
                    cmd.Parameters.AddWithValue("@EmpNo", DBNull.Value);
                else
                    cmd.Parameters.AddWithValue("@EmpNo", EmpNo.Text);
                cmd.Parameters.AddWithValue("@Name", NametextBox.Text);
                if (string.IsNullOrWhiteSpace(Designation.Text))
                    cmd.Parameters.AddWithValue("@Designation", DBNull.Value);
                else
                    cmd.Parameters.AddWithValue("@Designation", Designation.Text);
                if (string.IsNullOrWhiteSpace(Mobile.Text))
                    cmd.Parameters.AddWithValue("@Mobile", DBNull.Value);
                else
                    cmd.Parameters.AddWithValue("@Mobile", Mobile.Text);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception e1)
            {
                MessageBox.Show(e1.Message);
            }
        }
        private void Addbtn_Click(object sender, EventArgs e)
        {
            try
            {
                if (num == 0)
                {
                    Insertdb();
                    AIASLdataGrid();
                }
                else if (num == 1)
                {
                    Insertdb();
                    ATCdataGrid();
                }
                else if (num == 2)
                {
                    Insertdb();
                    BPCLdataGrid();
                }
                else if (num == 3)
                {
                    Insertdb();
                    CISFdataGrid();
                }
                else if (num == 4)
                {
                    Insertdb();
                    CNSITdataGrid();
                }
                else if (num == 5)
                {
                    Insertdb();
                    FIREGUJSAILdataGrid();
                }
                else if (num == 6)
                {
                    Insertdb();
                    MTdataGrid();
                }
                else
                {
                    MessageBox.Show("Else Condition is Excecuted.");
                }
            }
            catch (Exception e1)
            {
                con.Close();
                MessageBox.Show(e1.Message);
            }
        }
        private void dataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                ID.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                EmpNo.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                NametextBox.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
                Designation.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
                Mobile.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
            }
            catch (Exception)
            {

            }
        }
        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            dataGridView1.Rows[e.RowIndex].Cells[0].Value = (e.RowIndex + 1).ToString();
        }
        private void DataEntry_Load(object sender, EventArgs e)
        {
            if (frm.AdminMode)
            {
                Deletebtn.Enabled = true;
                Updatebtn.Enabled = true;
            }
            else
            {
                Deletebtn.Enabled = false;
                Updatebtn.Enabled = false;
            }
        }
    }
}
