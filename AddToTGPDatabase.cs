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
    public partial class AddToTGPDatabase : Form
    {
        string item = "";
        string itemType = "";

        public AddToTGPDatabase()
        {
            InitializeComponent();
        }

        public AddToTGPDatabase(string item_, string itemType_)
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
            public string tgp_srp;
            public string landed_cost;
            public double priceDiff;
        }
        List<ItemInfo> databaseItems = new List<ItemInfo>();

        public static string filePath_temp = System.Reflection.Assembly.GetExecutingAssembly().Location;
        public string dbFilePath = System.IO.Path.GetDirectoryName(filePath_temp) + "\\Database\\TGP_Database.xml";

        public void AddXmlNode(String sXml, String sNode, String sMenuNode, String sTypeAttrib_num, String sTypeAttrib_upc, String sTypeAttrib_desc, String sTypeAttrib_pk, String sTypeAttrib_tgp_srp, String sTypeAttrib_landed_cost)
        {
            XmlDocument xml;
            XmlElement xmlEle;
            XmlAttribute xmlAtb;
            XmlNode newNode;

            xml = new XmlDocument();
            xml.Load(dbFilePath);
            newNode = xml.SelectSingleNode(sNode);
            xmlEle = xml.CreateElement(sMenuNode);

            xmlAtb = xml.CreateAttribute("item_num");
            xmlAtb.Value = sTypeAttrib_num;
            xmlEle.SetAttributeNode(xmlAtb);

            xmlAtb = xml.CreateAttribute("upc");
            xmlAtb.Value = sTypeAttrib_upc;
            xmlEle.SetAttributeNode(xmlAtb);

            xmlAtb = xml.CreateAttribute("desc");
            xmlAtb.Value = sTypeAttrib_desc;
            xmlEle.SetAttributeNode(xmlAtb);

            xmlAtb = xml.CreateAttribute("pk");
            xmlAtb.Value = sTypeAttrib_pk;
            xmlEle.SetAttributeNode(xmlAtb);

            xmlAtb = xml.CreateAttribute("TGP_srp");
            xmlAtb.Value = sTypeAttrib_tgp_srp;
            xmlEle.SetAttributeNode(xmlAtb);

            xmlAtb = xml.CreateAttribute("Landed_cost");
            xmlAtb.Value = sTypeAttrib_landed_cost;
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
                temp.itemNum = xnl.Attributes["item_num"].Value;
                temp.itemUPC = xnl.Attributes["upc"].Value;
                temp.itemDesc = xnl.Attributes["desc"].Value;
                temp.itemPk = xnl.Attributes["pk"].Value;
                temp.tgp_srp = xnl.Attributes["TGP_srp"].Value;
                temp.landed_cost = xnl.Attributes["Landed_cost"].Value;

                databaseItems.Add(temp);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string itemNum = textBox1.Text;
            string upc = textBox2.Text;
            string desc = textBox3.Text;
            string pk = textBox4.Text;
            string tgp_srp = textBox5.Text;
            string landed_cost = textBox6.Text;

            if (itemNum == "" || upc == "" || desc == "" || pk == "" || tgp_srp == "" || landed_cost == "")
            {
                MessageBox.Show("Please enter all information.");
            }
            else
            {
                AddXmlNode(dbFilePath, "items", "itemInfo", itemNum, upc, desc, pk, tgp_srp, landed_cost);

                MessageBox.Show("The item is successfully added to TGP Database.", "Message Box");

                ReadDatabaseItems();
                this.Close();
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if(!(char.IsDigit(e.KeyChar) || e.KeyChar == Convert.ToChar(Keys.Back)))
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

        private void textBox6_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                this.button1_Click(sender, e);
            }
        }
    }
}
