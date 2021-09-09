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
    public partial class Form4 : Form
    {
        public Form4()
        {
            InitializeComponent();
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        // Cancel
        private void button2_Click(object sender, EventArgs e)
        {
            Form3.cancelChecker = true;
            this.Close();
        }

        // OK
        private void button1_Click(object sender, EventArgs e)
        {
            Form3 form3 = (Form3)Owner;
            form3.receivedData = textBox1.Text;

            if(form3.receivedData == "")
            {
                MessageBox.Show("Please Enter the quantity.", "Message Box");
            }
            else
            {
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
    }
}
