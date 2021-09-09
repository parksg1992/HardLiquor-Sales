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
    public partial class Edit : Form
    {
        public Edit()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // CANCEL
            Form3.cancelFlag = true;

            this.Close();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            // OK
            Form3.editQuantity = textBox1.Text;
            
            if (textBox1.Text == "")
            {
                MessageBox.Show("Please Enter The Quantity", "Message Box");
            }
            else
            {
                MessageBox.Show("Edit completed", "Message Box");
                this.Close();
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Size
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                button1_Click(sender, e);
            }
        }
    }
}
