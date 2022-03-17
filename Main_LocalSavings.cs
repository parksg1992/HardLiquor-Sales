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
    public partial class Main_LocalSavings : Form
    {
        public Main_LocalSavings()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            LocalSavings newForm = new LocalSavings();
            newForm.ShowDialog();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Search_LocalSavings newForm = new Search_LocalSavings();
            newForm.ShowDialog();
        }
    }
}
