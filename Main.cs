using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace HardLiquor_Sales
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form1 newForm = new Form1();
            newForm.ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form2 newForm = new Form2();
            newForm.ShowDialog();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string test1 = "Test1";
            string test2 = "Test2";
            
            using (StreamWriter wr = new StreamWriter("test.csv"))
            {
                wr.WriteLine("Article ID,Quantity");
                wr.WriteLine("{0},{1}", test1, test2);
                
                MessageBox.Show("File Create Test Complete", "Message Box");
            }

            File.Delete("test.csv");

            Form3 newForm = new Form3();
            newForm.ShowDialog();
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            LocalSaving localSaving = new LocalSaving();
            localSaving.ShowDialog();
        }
    }
}
