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
using System.Xml.Linq;

namespace HardLiquor_Sales
{
    public partial class Search_LocalSavings : Form
    {
        public static string filePath_temp = System.Reflection.Assembly.GetExecutingAssembly().Location;
        public string dbFilePath = System.IO.Path.GetDirectoryName(filePath_temp) + "\\LocalSavings_Database.xml";

        public struct ItemInfo
        {
            public string itemNum;
            public string itemUPC;
            public string itemDesc;
            public string itemPk;
            public string itemSave;
        }

        List<ItemInfo> databaseItems = new List<ItemInfo>();
        int getIndex = 0;

        public Search_LocalSavings()
        {
            InitializeComponent();
            checkedListBox1.SetItemChecked(0, true);

            ReadDatabaseItems();
        }

        // Read TGP Database.xml
        public void ReadDatabaseItems()
        {
            databaseItems.Clear();

            XmlDocument xml = new XmlDocument();
            xml.Load(dbFilePath);
            XmlNode xmlList_1 = xml.SelectNodes("items")[0];

            foreach (XmlNode xnl in xmlList_1)
            {
                var temp = new ItemInfo();
                temp.itemNum = xnl.Attributes["order_num"].Value;
                temp.itemUPC = xnl.Attributes["upc"].Value;
                temp.itemDesc = xnl.Attributes["name"].Value;
                temp.itemPk = xnl.Attributes["case"].Value;
                temp.itemSave = xnl.Attributes["save"].Value;

                databaseItems.Add(temp);
            }
        }

        private void checkedListBox1_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            for (int i = 0; i < checkedListBox1.Items.Count; ++i)
            {
                if (i != e.Index)
                {
                    checkedListBox1.SetItemChecked(i, false);
                }
            }
        }

        protected XmlNode CreateNode(XmlDocument xmlDoc, string name, string innerXml)
        {
            XmlNode node = xmlDoc.CreateElement(string.Empty, name, string.Empty);
            node.InnerXml = innerXml;

            return node;
        }
        public void textBoxEnable(bool enable)
        {
            if (enable == true)
            {
                textBox2.Enabled = true;
                textBox3.Enabled = true;
                textBox4.Enabled = true;
                textBox5.Enabled = true;
                textBox6.Enabled = true;
            }
            else
            {
                textBox2.Enabled = false;
                textBox3.Enabled = false;
                textBox4.Enabled = false;
                textBox5.Enabled = false;
                textBox6.Enabled = false;
            }
        }

        private void XMLModifier()
        {
            XmlDocument XmlDoc = new XmlDocument();
            XmlDoc.Load(dbFilePath);

            XmlNodeList nodeList = XmlDoc.SelectNodes("/descendant::items/itemInfo");

            // Attributes[0]; // Item Description
            // Attributes[1]; // Item Number
            // Attributes[2]; // Item UPC
            // Attributes[3]; // Item Pk
            // Attributes[4]; // Item Save

            nodeList[getIndex].Attributes[0].InnerText = textBox4.Text;
            nodeList[getIndex].Attributes[1].InnerText = textBox2.Text;
            nodeList[getIndex].Attributes[2].InnerText = textBox3.Text;
            nodeList[getIndex].Attributes[3].InnerText = textBox5.Text;
            nodeList[getIndex].Attributes[4].InnerText = textBox6.Text;

            XmlDoc.Save(dbFilePath);
            ReadDatabaseItems();

            MessageBox.Show("The item info is successfully updated.", "Message Box");

            textBoxEnable(false);
        }

        private void search_btn_Click(object sender, EventArgs e)
        {
            textBoxEnable(false);

            bool itemNumber = checkedListBox1.GetItemChecked(0);
            bool upc = checkedListBox1.GetItemChecked(1);

            ItemInfo item = new ItemInfo();

            // Found
            if (itemNumber == true || upc == true)
            {
                if (itemNumber == true)
                {
                    item = databaseItems.Find(x => x.itemNum == textBox1.Text);
                    getIndex = databaseItems.FindIndex(x => x.itemNum == textBox1.Text);

                    if (item.itemNum == null)
                    {
                        MessageBox.Show("The item is not in Database.", "Message Box");

                        AddToLocalSavingsDatabase newForm = new AddToLocalSavingsDatabase(textBox1.Text, "itemNum");
                        newForm.ShowDialog();

                        textBox1.Clear();
                    }
                }
                else
                {
                    item = databaseItems.Find(x => x.itemUPC == textBox1.Text);
                    getIndex = databaseItems.FindIndex(x => x.itemUPC == textBox1.Text);

                    if (item.itemUPC == null)
                    {
                        MessageBox.Show("The item is not in Database.", "Message Box");

                        AddToLocalSavingsDatabase newForm = new AddToLocalSavingsDatabase(textBox1.Text, "itemUPC");
                        newForm.ShowDialog();

                        textBox1.Clear();
                    }
                }

                textBox2.Text = item.itemNum;
                textBox3.Text = item.itemUPC;
                textBox4.Text = item.itemDesc;
                textBox5.Text = item.itemPk;
                textBox6.Text = item.itemSave;
            }
            else
            {
                MessageBox.Show("Please check one of the search options", "Message Box");
            }
        }

        private void edit_btn_Click(object sender, EventArgs e)
        {
            textBoxEnable(true);
        }

        private void save_btn_Click(object sender, EventArgs e)
        {
            XMLModifier();
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                this.search_btn_Click(sender, e);
            }
        }
    }
}
