using Nsoft;
using System;
using System.Data;
using System.Windows.Forms;

namespace Perde_Evim
{
    public partial class AdminPanel : Form
    {
        public AdminPanel()
        {
            InitializeComponent();
        }

        public void myRefresh()
        {
            MyData.selectCommand("Security", "SELECT * FROM Parol order by Status");
            MyData.dtmainParol = new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainParol);
            dataGridView1.DataSource = MyData.dtmainParol;
        }

        private void AdminPanel_Load(object sender, EventArgs e)
        {
            myRefresh();
        }

        private void DataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            MyData.updateCommand("Security", "UPDATE Parol SET "
                + "UserName='" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["UserName"].Value + "',"
                + "Parol='" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["Parol"].Value + "',"
                + "Status='" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["Status"].Value + "'"
                + " WHERE Код=" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["Код"].Value);

        }
    }
}
