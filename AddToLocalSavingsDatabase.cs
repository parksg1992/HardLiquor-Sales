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
    public partial class AddToLocalSavingsDatabase : Form
    {
        string item = "";
        string itemType = "";

        public AddToLocalSavingsDatabase()
        {
            InitializeComponent();
        }

        public AddToLocalSavingsDatabase(string item_, string itemType_)
        {
            InitializeComponent();
            item = item_;
            itemType = itemType_;

            if (itemType == "itemNum")
            {
                textBox1.Text = item;
            }
            else if (itemType == "itemUPC")
            {
                textBox2.Text = item;
            }
            else
            {

            }
        }

        public struct ItemInfo
        {
            public string itemNum;
            public string itemUPC;
            public string itemDesc;
            public string itemPk;
            public string itemSave;
        }
        List<ItemInfo> databaseItems = new List<ItemInfo>();

        public static string filePath_temp = System.Reflection.Assembly.GetExecutingAssembly().Location;
        public string dbFilePath = System.IO.Path.GetDirectoryName(filePath_temp) + "\\LocalSavings_Database.xml";

        public void AddXmlNode(String sXml, String sNode, String sMenuNode, String sTypeAttrib_num, String sTypeAttrib_upc, String sTypeAttrib_desc, String sTypeAttrib_pk, String sTypeAttrib_save)
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
            xmlAtb.Value = sTypeAttrib_desc;
            xmlEle.SetAttributeNode(xmlAtb);

            xmlAtb = xml.CreateAttribute("order_num");
            xmlAtb.Value = sTypeAttrib_num;
            xmlEle.SetAttributeNode(xmlAtb);

            xmlAtb = xml.CreateAttribute("upc");
            xmlAtb.Value = sTypeAttrib_upc;
            xmlEle.SetAttributeNode(xmlAtb);

            xmlAtb = xml.CreateAttribute("case");
            xmlAtb.Value = sTypeAttrib_pk;
            xmlEle.SetAttributeNode(xmlAtb);

            xmlAtb = xml.CreateAttribute("save");
            xmlAtb.Value = sTypeAttrib_save;
            xmlEle.SetAttributeNode(xmlAtb);

            newNode.AppendChild(xmlEle);
            xml.Save(dbFilePath);
            xml = null;
        }

        // Read TGP Database.xml
        public void ReadDatabaseItems()
        {
            XmlDocument xml = new XmlDocument();
            xml.Load(dbFilePath);
            XmlNode xmlList_1 = xml.SelectNodes("items")[0];

            foreach (XmlNode xnl in xmlList_1)
            {
                var temp = new ItemInfo();
                temp.itemDesc = xnl.Attributes["name"].Value;
                temp.itemNum = xnl.Attributes["order_num"].Value;
                temp.itemUPC = xnl.Attributes["upc"].Value;
                temp.itemPk = xnl.Attributes["case"].Value;
                temp.itemSave = xnl.Attributes["save"].Value;

                databaseItems.Add(temp);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            

            string itemNum = textBox1.Text;
            string upc = textBox2.Text;
            string desc = textBox3.Text;
            string pk = textBox4.Text;
            string save = textBox5.Text;

            if (itemNum == "" || upc == "" || desc == "" || pk == "" || save == "")
            {
                MessageBox.Show("Please enter all information.");
            }
            else
            {
                AddXmlNode(dbFilePath, "items", "itemInfo", itemNum, upc, desc, pk, save);

                MessageBox.Show("The item is successfully added to TGP Database.", "Message Box");

                ReadDatabaseItems();
                this.Close();
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(char.IsDigit(e.KeyChar) || e.KeyChar == Convert.ToChar(Keys.Back)))
            {
                e.Handled = true;
            }
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(char.IsDigit(e.KeyChar) || e.KeyChar == Convert.ToChar(Keys.Back)))
            {
                e.Handled = true;
            }
        }

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(char.IsDigit(e.KeyChar) || e.KeyChar == Convert.ToChar(Keys.Back)))
            {
                e.Handled = true;
            }
        }

        private void textBox5_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                this.button1_Click(sender, e);
            }
        }
    }
}
