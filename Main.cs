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
using Excel = Microsoft.Office.Interop.Excel;

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
            Form3 newForm = new Form3();
            newForm.ShowDialog();
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            Main_LocalSavings main_localSaving = new Main_LocalSavings();
            main_localSaving.ShowDialog();
        }


        /// //////////////////////////////////////////////////////////////////////

        Excel.Application application;
        Excel.Workbook workbook;
        Excel.Worksheet worksheet;
        object[,] itemNumRawData;
        object[,] itemUPCRawData;
        object[,] itemDescRawData;
        object[,] itemPkRawData;
        object[,] itemTgpSrpRawData;
        object[,] itemLandedCostRawData;

        OpenFileDialog OrderOFD = new OpenFileDialog();
        public static string TestFilePath = "";
        int orderFileCnt = 0;

        public struct ItemInfo_Test
        {
            public string itemNumber;
            public string itemUPC;
            public string itemDesc;
            public string itemPk;
            public string itemTgpSrp;
            public string itemLandedCost;

        }

        public string SetFilePath(bool isOrderFile)
        {
            if (isOrderFile == true)
            {
                return OrderOFD.FileName;
            }
            else
            {
                return "";
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (OrderOFD.ShowDialog() == DialogResult.OK)
            {
                TestFilePath = SetFilePath(true);
            }
            string destinationFile = TestFilePath;

            application = new Excel.Application();
            workbook = application.Workbooks.Open(Filename: destinationFile);
            worksheet = workbook.Worksheets.get_Item(1);
            application.Visible = false;

            orderFileCnt = worksheet.UsedRange.Rows.Count;

            // Item Number
            Excel.Range itemNumStartRange = worksheet.Cells[4, 6];
            Excel.Range itemNumEndRange = worksheet.Cells[orderFileCnt, 6];
            Excel.Range itemNumRange = worksheet.get_Range(itemNumStartRange, itemNumEndRange);
            itemNumRawData = itemNumRange.Value;

            // Item UPC
            Excel.Range itemUPCStartRange = worksheet.Cells[4, 7];
            Excel.Range itemUPCEndRange = worksheet.Cells[orderFileCnt, 7];
            Excel.Range itemUPCRange = worksheet.get_Range(itemUPCStartRange, itemUPCEndRange);
            itemUPCRawData = itemUPCRange.Value;

            // Item Desc
            Excel.Range itemDescStartRange = worksheet.Cells[4, 8];
            Excel.Range itemDescEndRange = worksheet.Cells[orderFileCnt, 8];
            Excel.Range itemDescRange = worksheet.get_Range(itemDescStartRange, itemDescEndRange);
            itemDescRawData = itemDescRange.Value;

            // Item PK
            Excel.Range itemPkStartRange = worksheet.Cells[4, 9];
            Excel.Range itemPkEndRange = worksheet.Cells[orderFileCnt, 9];
            Excel.Range itemPkRange = worksheet.get_Range(itemPkStartRange, itemPkEndRange);
            itemPkRawData = itemPkRange.Value;

            // Item TGP SRP
            Excel.Range itemTgpSrpStartRange = worksheet.Cells[4, 14];
            Excel.Range itemTgpSrpEndRange = worksheet.Cells[orderFileCnt, 14];
            Excel.Range itemTgpSrpRange = worksheet.get_Range(itemTgpSrpStartRange, itemTgpSrpEndRange);
            itemTgpSrpRawData = itemTgpSrpRange.Value;

            // Item Landed Cost
            Excel.Range itemLandedCostStartRange = worksheet.Cells[4, 16];
            Excel.Range itemLandedCostEndRange = worksheet.Cells[orderFileCnt, 16];
            Excel.Range itemLandedCostRange = worksheet.get_Range(itemLandedCostStartRange, itemLandedCostEndRange);
            itemLandedCostRawData = itemLandedCostRange.Value;


            string savePath = @"C:\Users\rhehf\Downloads\test.xml";

            for (int i = 1; i <= itemNumRawData.GetLength(0); i++)
            {
                string itemInfoText = "<itemInfo item_num= \"" + itemNumRawData[i, 1].ToString() + "\"" + " upc = \"" + itemUPCRawData[i, 1].ToString() + "\"" +
                    " desc = \"" + itemDescRawData[i,1].ToString() + "\"" + " pk = \"" + itemPkRawData[i,1].ToString() + "\"" + " TGP_srp = \""+ itemTgpSrpRawData[i,1].ToString() +
                   "\"" + " Landed_cost = \"" + itemLandedCostRawData[i,1].ToString() + "\"" + "/>" + "\r\n";
                System.IO.File.AppendAllText(savePath, itemInfoText, Encoding.Default);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Form6 newForm = new Form6();
            newForm.ShowDialog();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            AddToTGPDatabase newForm = new AddToTGPDatabase();
            newForm.ShowDialog();
        }
    }
}