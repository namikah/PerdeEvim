using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Xml;
using System.Data.SqlClient;
using System.IO;
using System.Globalization;
using Nsoft;

namespace Perde_Evim
{
    public partial class Parol : Form
    {
        public Parol()
        {
            InitializeComponent();
        }

        private void TxtParolGiris_Click(object sender, EventArgs e)
        {
            MyData.selectCommand("Security", "SELECT * FROM Parol WHERE UserName='" + Environment.UserName + "'");
            MyData.dtmainParol = new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainParol);

            if (MyData.dtmainParol.Rows[0]["UserName"].ToString() == Environment.UserName && MyData.dtmainParol.Rows[0]["Parol"].ToString() == txtParol.Text)
            {
                MyCheck.Parolicaze = true;
                base.Close();
            }
        }

        private void TxtParol_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                TxtParolGiris_Click(sender, e);
            }
        }
    }
}
