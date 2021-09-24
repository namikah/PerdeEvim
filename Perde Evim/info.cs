using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Perde_Evim
{
    public partial class info : Form
    {
        public info()
        {
            InitializeComponent();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (button1.Visible == false)
            {
                button1.Visible = true;
                return;
            }
            else { button1.Visible = false; timer2.Enabled = true; timer1.Enabled = false; }
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            if (button2.Visible == false)
            {
                button2.Visible = true;
                return;
            }
            else { button2.Visible = false; timer1.Enabled = true; timer2.Enabled = false; }
        }

        private void Info_Load(object sender, EventArgs e)
        {
            base.Text = "Version: " + ProductVersion;
        }
    }
}
