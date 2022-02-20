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
    public partial class Form5 : Form
    {
        public struct ItemInfo
        {
            public string itemNum;
            public string itemUPC;
            public string itemDesc;
            public string itemPk;
            public string tgp_srp;
            public string landed_cost;
        }
        List<ItemInfo> databaseItems = new List<ItemInfo>();

        public Form5()
        {
            InitializeComponent();
            Read();
        }
        
        OpenFileDialog TGP_OFD = new OpenFileDialog();
        public static string filePath_temp = System.Reflection.Assembly.GetExecutingAssembly().Location;
        public string dbFilePath = System.IO.Path.GetDirectoryName(filePath_temp) + "\\TGP_Database.xml";
        public string TGPPriceCheckerExcelFile = System.IO.Path.GetDirectoryName(filePath_temp) + "\\TGP_PriceChecker_";

        public void Read()
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
            if (TGP_OFD.ShowDialog() == DialogResult.OK)
            {
                textBox1.Clear();
                textBox1.Text = TGP_OFD.FileName;
                //OrderFilePath = SetFilePath(true);
            }
        }
    }
}
