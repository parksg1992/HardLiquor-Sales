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
    public partial class Form3 : Form
    {
        public struct ItemOrderForm
        {
            public string IOF_productType;
            public string IOF_productName;
            public string IOF_productSize;
            public string IOF_productNumber;
            public string IOF_productQuantity;
            public bool IOF_productAdded;
        }

        public struct ItemInfo
        {
            public string productNumber;
            public string productName;
            public string productSize;
        }

        public static bool cancelChecker = false;
        public static List<ItemOrderForm> orderList = new List<ItemOrderForm>();
        //public string filePath = "C:\\Users\\rhehf\\OneDrive\\Desktop\\database.xml";
        public static string filePath_temp = System.Reflection.Assembly.GetExecutingAssembly().Location;
        public string dbFilePath = System.IO.Path.GetDirectoryName(filePath_temp) + "\\Database\\Liquor_Database.xml";
        public string csvFilePath = System.IO.Path.GetDirectoryName(filePath_temp);
        public string orderFileListPath = System.IO.Path.GetDirectoryName(filePath_temp) + "\\Order List.txt";

        //        static string year_log = DateTime.Now.ToString("yyyy");
        //        static string month_log = DateTime.Now.ToString("MM");
        //        static string date_log = DateTime.Now.ToString("dd");
        //        public string logFilePath = System.IO.Path.GetDirectoryName(filePath_temp) + "\\Log" + "\\" + year_log + "\\" + month_log + "\\" + date_log;

        public static string editQuantity;
        string date = DateTime.Now.ToString("yyyy-MM-dd__hh-mm-ss");

        List<ItemInfo> _6b = new List<ItemInfo>();
        List<ItemInfo> _8b = new List<ItemInfo>();
        List<ItemInfo> _12b = new List<ItemInfo>();
        List<ItemInfo> _18b = new List<ItemInfo>();
        List<ItemInfo> _24b = new List<ItemInfo>();
        List<ItemInfo> _1c = new List<ItemInfo>();
        List<ItemInfo> _4c = new List<ItemInfo>();
        List<ItemInfo> _6c = new List<ItemInfo>();
        List<ItemInfo> _8c = new List<ItemInfo>();
        List<ItemInfo> _12c = new List<ItemInfo>();
        List<ItemInfo> _15c = new List<ItemInfo>();
        List<ItemInfo> _18c = new List<ItemInfo>();
        List<ItemInfo> _24c = new List<ItemInfo>();
        List<ItemInfo> _30c = new List<ItemInfo>();
        List<ItemInfo> _36c = new List<ItemInfo>();
        List<ItemInfo> _etc = new List<ItemInfo>();

        public Form3()
        {
            InitializeComponent();
        }

        public string receivedData;
        int itemIndex = 0;
        public static List<ItemOrderForm> finalIOF = new List<ItemOrderForm>();

        public void ClearList()
        {
            _6b.Clear();
            _8b.Clear();
            _12b.Clear();
            _18b.Clear();
            _24b.Clear();
            _1c.Clear();
            _4c.Clear();
            _6c.Clear();
            _8c.Clear();
            _12c.Clear();
            _15c.Clear();
            _18c.Clear();
            _24c.Clear();
            _30c.Clear();
            _36c.Clear();
            _etc.Clear();
        }

        public string GetSelectedItem_listbox1(int selectedItemIdx)
        {
            string selectedItem = "";

            switch (selectedItemIdx)
            {
                case 0:
                    selectedItem = "_6B";
                    itemIndex = 0;
                    break;
                case 1:
                    selectedItem = "_8B";
                    itemIndex = 1;
                    break;
                case 2:
                    selectedItem = "_12B";
                    itemIndex = 2;
                    break;
                case 3:
                    selectedItem = "_18B";
                    itemIndex = 3;
                    break;
                case 4:
                    selectedItem = "_24B";
                    itemIndex = 4;
                    break;
                case 5:
                    selectedItem = "_1C";
                    itemIndex = 5;
                    break;
                case 6:
                    selectedItem = "_4C";
                    itemIndex = 6;
                    break;
                case 7:
                    selectedItem = "_6C";
                    itemIndex = 7;
                    break;
                case 8:
                    selectedItem = "_8C";
                    itemIndex = 8;
                    break;
                case 9:
                    selectedItem = "_12C";
                    itemIndex = 9;
                    break;
                case 10:
                    selectedItem = "_15C";
                    itemIndex = 10;
                    break;
                case 11:
                    selectedItem = "_18C";
                    itemIndex = 11;
                    break;
                case 12:
                    selectedItem = "_24C";
                    itemIndex = 12;
                    break;
                case 13:
                    selectedItem = "_30C";
                    itemIndex = 13;
                    break;
                case 14:
                    selectedItem = "_36C";
                    itemIndex = 14;
                    break;
                case 15:
                    selectedItem = "_ETC";
                    itemIndex = 15;
                    break;
                default:
                    break;
            }
            return selectedItem;
        }

        public void Read()
        {
            ClearList();
            //string temp = "";
            XmlDocument xml = new XmlDocument();
            xml.Load(dbFilePath);
            XmlNode xmlList_1 = xml.SelectNodes("beer")[0];

            var selectedItemIndex = listBox1.SelectedIndex;
            string selectedItem = GetSelectedItem_listbox1(selectedItemIndex);

            foreach (XmlNode xnl in xmlList_1.SelectNodes(selectedItem))
            {
                var temp = new ItemInfo();
                temp.productName = xnl.Attributes["name"].Value;
                temp.productNumber = xnl.Attributes["num"].Value;
                temp.productSize = xnl.Attributes["size"].Value;

                switch (itemIndex)
                {
                    case 0:
                        _6b.Add(temp);
                        break;
                    case 1:
                        _8b.Add(temp);
                        break;
                    case 2:
                        _12b.Add(temp);
                        break;
                    case 3:
                        _18b.Add(temp);
                        break;
                    case 4:
                        _24b.Add(temp);
                        break;
                    case 5:
                        _1c.Add(temp);
                        break;
                    case 6:
                        _4c.Add(temp);
                        break;
                    case 7:
                        _6c.Add(temp);
                        break;
                    case 8:
                        _8c.Add(temp);
                        break;
                    case 9:
                        _12c.Add(temp);
                        break;
                    case 10:
                        _15c.Add(temp);
                        break;
                    case 11:
                        _18c.Add(temp);
                        break;
                    case 12:
                        _24c.Add(temp);
                        break;
                    case 13:
                        _30c.Add(temp);
                        break;
                    case 14:
                        _36c.Add(temp);
                        break;
                    case 15:
                        _etc.Add(temp);
                        break;
                    default:
                        break;
                }
            }
        }
        public void LoadItemList()
        {
            Read();

            // Clear all items in listbox
            listBox2.Items.Clear();

            // 6B
            if (listBox1.SelectedIndex == 0)
            {
                for (int idx = 0; idx < _6b.Count; idx++)
                {
                    listBox2.Items.Add(_6b[idx].productName + ",  " + _6b[idx].productSize);
                }
            }

            // 8B
            if (listBox1.SelectedIndex == 1)
            {
                for (int idx = 0; idx < _8b.Count; idx++)
                {
                    listBox2.Items.Add(_8b[idx].productName + ",  " + _8b[idx].productSize);
                }
            }

            // 12B
            if (listBox1.SelectedIndex == 2)
            {
                for (int idx = 0; idx < _12b.Count; idx++)
                {
                    listBox2.Items.Add(_12b[idx].productName + ",  " + _12b[idx].productSize);
                }
            }

            // 18B
            if (listBox1.SelectedIndex == 3)
            {
                for (int idx = 0; idx < _18b.Count; idx++)
                {
                    listBox2.Items.Add(_18b[idx].productName + ",  " + _18b[idx].productSize);
                }
            }

            // 24B
            if (listBox1.SelectedIndex == 4)
            {
                for (int idx = 0; idx < _24b.Count; idx++)
                {
                    listBox2.Items.Add(_24b[idx].productName + ",  " + _24b[idx].productSize);
                }
            }

            // 1c
            if (listBox1.SelectedIndex == 5)
            {
                for (int idx = 0; idx < _1c.Count; idx++)
                {
                    listBox2.Items.Add(_1c[idx].productName + ",  " + _1c[idx].productSize);
                }
            }

            // 4c
            if (listBox1.SelectedIndex == 6)
            {
                for (int idx = 0; idx < _4c.Count; idx++)
                {
                    listBox2.Items.Add(_4c[idx].productName + ",  " + _4c[idx].productSize);
                }
            }

            // 6c
            if (listBox1.SelectedIndex == 7)
            {
                for (int idx = 0; idx < _6c.Count; idx++)
                {
                    listBox2.Items.Add(_6c[idx].productName + ",  " + _6c[idx].productSize);
                }
            }

            // 8c
            if (listBox1.SelectedIndex == 8)
            {
                for (int idx = 0; idx < _8c.Count; idx++)
                {
                    listBox2.Items.Add(_8c[idx].productName + ",  " + _8c[idx].productSize);
                }
            }

            // 12c
            if (listBox1.SelectedIndex == 9)
            {
                for (int idx = 0; idx < _12c.Count; idx++)
                {
                    listBox2.Items.Add(_12c[idx].productName + ",  " + _12c[idx].productSize);
                }
            }

            // 15c
            if (listBox1.SelectedIndex == 10)
            {
                for (int idx = 0; idx < _15c.Count; idx++)
                {
                    listBox2.Items.Add(_15c[idx].productName + ",  " + _15c[idx].productSize);
                }
            }

            // 18c
            if (listBox1.SelectedIndex == 11)
            {
                for (int idx = 0; idx < _18c.Count; idx++)
                {
                    listBox2.Items.Add(_18c[idx].productName + ",  " + _18c[idx].productSize);
                }
            }

            // 24c
            if (listBox1.SelectedIndex == 12)
            {
                for (int idx = 0; idx < _24c.Count; idx++)
                {
                    listBox2.Items.Add(_24c[idx].productName + ",  " + _24c[idx].productSize);
                }
            }

            // 30c
            if (listBox1.SelectedIndex == 13)
            {
                for (int idx = 0; idx < _30c.Count; idx++)
                {
                    listBox2.Items.Add(_30c[idx].productName + ",  " + _30c[idx].productSize);
                }
            }

            // 36c
            if (listBox1.SelectedIndex == 14)
            {
                for (int idx = 0; idx < _36c.Count; idx++)
                {
                    listBox2.Items.Add(_36c[idx].productName + ",  " + _36c[idx].productSize);
                }
            }

            // 4c
            if (listBox1.SelectedIndex == 15)
            {
                for (int idx = 0; idx < _etc.Count; idx++)
                {
                    listBox2.Items.Add(_etc[idx].productName + ",  " + _etc[idx].productSize);
                }
            }
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadItemList();
        }

        private void listBox2_MouseClick(object sender, MouseEventArgs e)
        {
            int selectedIndex = -1;
            Point point = e.Location;

            selectedIndex = listBox2.IndexFromPoint(point);

            ItemOrderForm iOF = new ItemOrderForm();

            if (selectedIndex != -1)
            {
                List<ItemInfo> listForNumber = new List<ItemInfo>();
                string temp = listBox1.SelectedItem.ToString();

                // Check if the selected item is already on the list
                bool isOnTheList = false;

                foreach (var data in finalIOF)
                {
                    if (data.IOF_productAdded == true && (listBox2.Text == data.IOF_productName)
                        && temp == data.IOF_productType)
                    {
                        MessageBox.Show("The selected item is already on the list.", "Message Box");
                        Log2FileClass.Log2File("ItemDoubleSelectCheck.log", "========== Item Double Selected  ========== \n"
                            + "OrderList Item : " + listBox2.Text
                            + ", Selected Item :" + data.IOF_productName + "\n"
                            + "OrderList Item Type : " + temp
                            + ", Selected Item Type : " + data.IOF_productType +"\n");

                        isOnTheList = true;
                    }
                }

                if (isOnTheList != true)
                {
                    Form4 newForm = new Form4();
                    newForm.Owner = this;
                    newForm.ShowDialog();

                    if (temp == "6B")
                    {
                        listForNumber = _6b;
                    }
                    else if (temp == "8B")
                    {
                        listForNumber = _8b;
                    }
                    else if (temp == "12B")
                    {
                        listForNumber = _12b;
                    }
                    else if (temp == "18B")
                    {
                        listForNumber = _18b;
                    }
                    else if (temp == "24B")
                    {
                        listForNumber = _24b;
                    }
                    else if (temp == "Single Can")
                    {
                        listForNumber = _1c;
                    }
                    else if (temp == "4C")
                    {
                        listForNumber = _4c;
                    }
                    else if (temp == "6C")
                    {
                        listForNumber = _6c;
                    }
                    else if (temp == "8C")
                    {
                        listForNumber = _8c;
                    }
                    else if (temp == "12C")
                    {
                        listForNumber = _12c;
                    }
                    else if (temp == "15C")
                    {
                        listForNumber = _15c;
                    }
                    else if (temp == "18C")
                    {
                        listForNumber = _18c;
                    }
                    else if (temp == "24C")
                    {
                        listForNumber = _24c;
                    }
                    else if (temp == "30C")
                    {
                        listForNumber = _30c;
                    }
                    else if (temp == "36C")
                    {
                        listForNumber = _36c;
                    }
                    else if (temp == "ETC")
                    {
                        listForNumber = _etc;
                    }
                    else
                    {

                    }
                    
                    if (cancelChecker == false && receivedData != "")
                    {
                        iOF.IOF_productName = listBox2.Text;
                        iOF.IOF_productType = temp;
                        iOF.IOF_productQuantity = receivedData;
                        iOF.IOF_productNumber = listForNumber[listBox2.SelectedIndex].productNumber;
                        iOF.IOF_productAdded = true;

                        finalIOF.Add(iOF);

                        listBox3.Items.Add(iOF.IOF_productType + " - " + iOF.IOF_productName + iOF.IOF_productSize + " : " + iOF.IOF_productQuantity);
                        listBox3.TopIndex = listBox3.Items.Count - 1;
                    }
                    cancelChecker = false;
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (finalIOF.Count == 0)
            {
                MessageBox.Show("There is nothing to be done.", "Message Box");
            }
            else
            {
                string csvFileFullPath = csvFilePath + "\\BDL Orderform_" + date;

                if (File.Exists(csvFileFullPath + ".csv"))
                {
                    File.Delete(csvFileFullPath + ".csv");
                }

                using (StreamWriter wr = new StreamWriter(csvFileFullPath + ".csv"))
                {
                    wr.WriteLine("Article ID,Quantity");

                    foreach (var data in finalIOF)
                    {
                        wr.WriteLine("{0},{1}", data.IOF_productNumber, data.IOF_productQuantity);

                    }

                    MessageBox.Show("Done", "Message Box");
                    this.Close();
                }

                if (File.Exists(orderFileListPath))
                {
                    File.Delete(orderFileListPath);
                }

                int line = 0;

                foreach (var data in finalIOF)
                {
                    line++;
                    if (data.IOF_productNumber == "0")
                    {
                        System.IO.File.AppendAllText(orderFileListPath, "**");
                    }
                    System.IO.File.AppendAllText(orderFileListPath, line + " > " + data.IOF_productType + " - " + data.IOF_productName + " " + data.IOF_productSize + " : " + data.IOF_productQuantity + "\n");
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Add to Database
            AddToDatabase newForm = new AddToDatabase();
            newForm.ShowDialog();
        }

        public void DelXmlNode(String sXml, String sNode)
        {
            XmlDocument xmlDoc;
            XmlNode parentNode, delNode;

            xmlDoc = new XmlDocument();
            xmlDoc.Load(dbFilePath);

            delNode = xmlDoc.SelectSingleNode(sNode);
            parentNode = delNode.ParentNode;

            parentNode.RemoveChild(delNode);

            xmlDoc.Save(dbFilePath);
            xmlDoc = null;
        }

        public void TestAttributes_Remove()
        {
            XmlDocument xml = new XmlDocument();
            xml.Load(dbFilePath);
            XmlNode xmlList_1 = xml.SelectNodes("beer")[0];

            XmlAttributeCollection col = xmlList_1.Attributes;
            XmlAttribute attr = col["_6B"];
            col.Remove(attr);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // REMOVE

            TestAttributes_Remove();

        }

        public void ShowList()
        {
            foreach (var data in Form3.finalIOF)
            {
                listBox3.Items.Add(data.IOF_productType + " - " + data.IOF_productName + data.IOF_productSize + " : " + data.IOF_productQuantity);
            }
        }

        public static bool cancelFlag = false;

        private void button3_Click(object sender, EventArgs e)
        {
            Edit edit = new Edit();
            edit.ShowDialog();

            if (cancelFlag == false && editQuantity != "")
            {
                int selectedIdx = listBox3.SelectedIndex;

                Form3.ItemOrderForm temp = Form3.finalIOF[selectedIdx];

                temp.IOF_productQuantity = Form3.editQuantity;
                temp.IOF_productName = Form3.finalIOF[selectedIdx].IOF_productName;
                temp.IOF_productNumber = Form3.finalIOF[selectedIdx].IOF_productNumber;
                temp.IOF_productSize = Form3.finalIOF[selectedIdx].IOF_productSize;
                temp.IOF_productType = Form3.finalIOF[selectedIdx].IOF_productType;

                Form3.finalIOF[selectedIdx] = temp;

                Log2FileClass.Log2File("EditFunction.log", "========== Edit Function  ========== \n"
                    + "CancelFlag : " + cancelFlag + "\n"
                    + ", Edit ProductQuantity : " + temp.IOF_productQuantity + "\n"
                    + ", Edit ProductName & Size : " + temp.IOF_productName + "\n"
                    + ", Edit ProductNumber : " + temp.IOF_productNumber + "\n"
                    + ", Edit ProductType : " + temp.IOF_productType + "\n");
            }

            cancelFlag = false;

            listBox3.Items.Clear();
            ShowList();
        }
    }
}