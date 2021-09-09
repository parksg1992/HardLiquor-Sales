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
    public partial class MoreFileToAdd : Form
    {
        public MoreFileToAdd()
        {
            InitializeComponent();
        }

        // NO
        private void button2_Click(object sender, EventArgs e)
        {
            LocalSaving.noMoreFileToAddFlag = true;
            this.Close();
        }

        // YES
        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
