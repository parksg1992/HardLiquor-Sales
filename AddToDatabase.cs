using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace HardLiquor_Sales
{
    public partial class AddToDatabase : Form
    {
        public AddToDatabase()
        {
            InitializeComponent();
        }
        public string size;
        public string num;
        public string name;
        public string type;
        public static string filePath_temp = System.Reflection.Assembly.GetExecutingAssembly().Location;
        public string dbFilePath = System.IO.Path.GetDirectoryName(filePath_temp) + "\\database.xml";
        //public string dbFilePath = "C:\\Users\\rhehf\\OneDrive\\Desktop\\database.xml";

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Order Num
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Size
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        public void AddXmlNode(String sXml, String sNode, String sMenuNode, String sTypeAttrib_name, String sTypeAttrib_num, String sTypeAttrib_size)
        {
            XmlDocument xml;
            XmlElement xmlEle;
            XmlAttribute xmlAtb;
            XmlNode newNode;
            
            xml = new XmlDocument();
            xml.Load(dbFilePath);
            newNode = xml.SelectSingleNode(sNode);
            xmlEle = xml.CreateElement(sMenuNode);
            xmlAtb = xml.CreateAttribute("name");
            xmlAtb.Value = sTypeAttrib_name;
            xmlEle.SetAttributeNode(xmlAtb);

            xmlAtb = xml.CreateAttribute("num");
            xmlAtb.Value = sTypeAttrib_num;
            xmlEle.SetAttributeNode(xmlAtb);

            xmlAtb = xml.CreateAttribute("size");
            xmlAtb.Value = sTypeAttrib_size + "ml";
            xmlEle.SetAttributeNode(xmlAtb);

            newNode.AppendChild(xmlEle);
            xml.Save(dbFilePath);
            xml = null;

            MessageBox.Show("Done", "Message Box");
            
        }
        
        private void button1_Click(object sender, EventArgs e)
        {
            if(listBox1.SelectedItems.Count == 0)
            {
                MessageBox.Show("Please Select Type.", "Message Box");

                return;
            }
            
            AddXmlNode(dbFilePath, "beer", type, textBox1.Text, textBox2.Text, textBox3.Text);

            textBox1.Clear();
            textBox2.Clear();
            textBox3.Clear();

            this.Close();
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            // 6B
            if (listBox1.SelectedIndex == 0)
            {
                type = "_6B";
            }
            // 8B
            else if (listBox1.SelectedIndex == 1)
            {
                type = "_8B";
            }
            // 12B
            else if (listBox1.SelectedIndex == 2)
            {
                type = "_12B";
            }
            // 18B
            else if (listBox1.SelectedIndex == 3)
            {
                type = "_18B";
            }
            // 24B
            else if (listBox1.SelectedIndex == 4)
            {
                type = "_24B";
            }
            // 1C
            else if (listBox1.SelectedIndex == 5)
            {
                type = "_1C";
            }
            // 4C
            else if (listBox1.SelectedIndex == 6)
            {
                type = "_4C";
            }
            // 6C
            else if (listBox1.SelectedIndex == 7)
            {
                type = "_6C";
            }
            // 8C
            else if (listBox1.SelectedIndex == 8)
            {
                type = "_8C";
            }
            // 12C
            else if (listBox1.SelectedIndex == 9)
            {
                type = "_12C";
            }
            // 15C
            else if (listBox1.SelectedIndex == 10)
            {
                type = "_15C";
            }
            // 18C
            else if (listBox1.SelectedIndex == 11)
            {
                type = "_18C";
            }
            // 24C
            else if (listBox1.SelectedIndex == 12)
            {
                type = "_24C";
            }
            // 30C
            else if (listBox1.SelectedIndex == 13)
            {
                type = "_30C";
            }
            // 36C
            else if (listBox1.SelectedIndex == 14)
            {
                type = "_36C";
            }
            // ETC
            else if (listBox1.SelectedIndex == 15)
            {
                type = "_ETC";
            }
            else
            {
                // Do Nothing
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void textBox3_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                button1_Click(sender, e);
            }
        }

        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                button1_Click(sender, e);
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
