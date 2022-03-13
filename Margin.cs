using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace HardLiquor_Sales
{
    public partial class Margin : Form
    {
        public Margin()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form5 form5 = (Form5)Owner;
            string text = textBox1.Text;

            if(text.Length > 2 || text.Length == 0)
            {
                MessageBox.Show("Margin should be less than 100%.\nPlease enter again.");
            }
            else
            {
                form5.receivedMarginData = Convert.ToInt32(text);

                this.DialogResult = DialogResult.OK;
                this.Close();
            }

        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                this.button1_Click(sender, e);
            }
        }

        private void textBox1_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }
    }
}
