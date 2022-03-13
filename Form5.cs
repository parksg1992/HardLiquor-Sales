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
using Excel = Microsoft.Office.Interop.Excel;

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
            public double priceDiff;
        }
        List<ItemInfo> databaseItems = new List<ItemInfo>();
        List<ItemInfo> tgpItems = new List<ItemInfo>();
        List<ItemInfo> updateItemList = new List<ItemInfo>();

        bool updateFlag = true;

        OpenFileDialog OrderOFD = new OpenFileDialog();

        public static string tgpExcelFilePath = "";
        int orderFileCnt = 0;

        object[,] itemNumRawData;
        object[,] itemUPCRawData;
        object[,] itemDescRawData;
        object[,] itemPkRawData;
        object[,] itemTgpSrpRawData;
        object[,] itemLandedCostRawData;

        Excel.Application application;
        Excel.Workbook workbook;
        Excel.Worksheet worksheet;

        public Form5()
        {
            InitializeComponent();
            ReadDatabaseItems();
        }

        OpenFileDialog TGP_OFD = new OpenFileDialog();
        public static string filePath_temp = System.Reflection.Assembly.GetExecutingAssembly().Location;
        public string dbFilePath = System.IO.Path.GetDirectoryName(filePath_temp) + "\\TGP_Database.xml";
        public string tgpExcelFile = System.IO.Path.GetDirectoryName(filePath_temp) + "\\TGP_Price_Checker_";
        public string TGPPriceCheckerExcelFile = System.IO.Path.GetDirectoryName(filePath_temp) + "\\TGP_PriceChecker_";


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

        int excelItemNum = 0;

        // Read TGP Item Excel file
        public void ReadTGPItems()
        {
            if (OrderOFD.ShowDialog() == DialogResult.OK)
            {
                textBox1.Clear();
                textBox1.Text = TGP_OFD.FileName;

                tgpExcelFilePath = SetFilePath(true);
            }
            string destinationFile = tgpExcelFilePath;

            pBar1.Value = 40;

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

            pBar1.Value = 60;

            // Item UPC
            Excel.Range itemUPCStartRange = worksheet.Cells[4, 7];
            Excel.Range itemUPCEndRange = worksheet.Cells[orderFileCnt, 7];
            Excel.Range itemUPCRange = worksheet.get_Range(itemUPCStartRange, itemUPCEndRange);
            itemUPCRawData = itemUPCRange.Value;

            pBar1.Value = 70;

            // Item Desc
            Excel.Range itemDescStartRange = worksheet.Cells[4, 8];
            Excel.Range itemDescEndRange = worksheet.Cells[orderFileCnt, 8];
            Excel.Range itemDescRange = worksheet.get_Range(itemDescStartRange, itemDescEndRange);
            itemDescRawData = itemDescRange.Value;

            pBar1.Value = 80;

            // Item PK
            Excel.Range itemPkStartRange = worksheet.Cells[4, 9];
            Excel.Range itemPkEndRange = worksheet.Cells[orderFileCnt, 9];
            Excel.Range itemPkRange = worksheet.get_Range(itemPkStartRange, itemPkEndRange);
            itemPkRawData = itemPkRange.Value;

            pBar1.Value = 90;

            // Item TGP SRP
            Excel.Range itemTgpSrpStartRange = worksheet.Cells[4, 14];
            Excel.Range itemTgpSrpEndRange = worksheet.Cells[orderFileCnt, 14];
            Excel.Range itemTgpSrpRange = worksheet.get_Range(itemTgpSrpStartRange, itemTgpSrpEndRange);
            itemTgpSrpRawData = itemTgpSrpRange.Value;

            pBar1.Value = 100;

            // Item Landed Cost
            Excel.Range itemLandedCostStartRange = worksheet.Cells[4, 16];
            Excel.Range itemLandedCostEndRange = worksheet.Cells[orderFileCnt, 16];
            Excel.Range itemLandedCostRange = worksheet.get_Range(itemLandedCostStartRange, itemLandedCostEndRange);
            itemLandedCostRawData = itemLandedCostRange.Value;

            pBar1.Value = 110;

            excelItemNum = itemNumRawData.GetLength(0);
        }

        public void AddToList()
        {
            for(int i = 1; i < excelItemNum; i++)
            {
                var temp = new ItemInfo();

                temp.itemNum = (itemNumRawData[i, 1] == null ? "0" : itemNumRawData[i, 1].ToString());
                temp.itemUPC = (itemUPCRawData[i, 1] == null ? "0" : itemUPCRawData[i, 1].ToString());
                temp.itemDesc = (itemDescRawData[i, 1] == null ? " " : itemDescRawData[i, 1].ToString());
                temp.itemPk = (itemPkRawData[i, 1] == null ? "0" : itemPkRawData[i, 1].ToString());
                temp.tgp_srp = (itemTgpSrpRawData[i, 1] == null ? "0" : itemTgpSrpRawData[i, 1].ToString());
                temp.landed_cost = (itemLandedCostRawData[i, 1] == null ? "0" : itemLandedCostRawData[i, 1].ToString());

                tgpItems.Add(temp);

                //pBar1.PerformStep();
            }
        }

        public int receivedMarginData = 0;

        public void Compare()
        {
            for (int i=0;i<databaseItems.Count(); i++)
            {
                double itemNum_DB = Convert.ToDouble(databaseItems[i].itemNum);

                var matchedItemList = tgpItems.Find(x => Convert.ToDouble(x.itemNum) == itemNum_DB);
                int matchedItemIndex = tgpItems.FindIndex(x => x.itemNum.Equals(matchedItemList.itemNum));

                if (matchedItemIndex != -1)
                {
                    double convertedItemNum_DB = Convert.ToDouble(matchedItemList.itemNum);
                    double convertedLandedCost_DB = Convert.ToDouble(databaseItems[i].landed_cost);
                    double convertedLandedCost_TGP = Convert.ToDouble(tgpItems[matchedItemIndex].landed_cost);

                    // FOUND
                    // Compare the landed cost
                    if(convertedLandedCost_TGP > convertedLandedCost_DB)
                    {
                        // Price Increased
                        matchedItemList.priceDiff = convertedLandedCost_TGP - convertedLandedCost_DB;
                        updateItemList.Add(matchedItemList);
                    }
                    else if(convertedLandedCost_TGP < convertedLandedCost_DB)
                    {
                        // Price Decreased
                        matchedItemList.priceDiff = convertedLandedCost_TGP - convertedLandedCost_DB;
                        updateItemList.Add(matchedItemList);
                    }
                }
                int a = pBar1.Value;

                if(pBar1.Value < ((pBar1.Maximum) * 0.8))
                {
                    pBar1.PerformStep();
                }
            }

            if(updateItemList.Count() != 0)
            {
                int updateItemNum = updateItemList.Count();

                MessageBox.Show(updateItemNum + " items need to be updated.\nPlease enter margin follow.");

                Margin newForm = new Margin();
                newForm.Owner = this;
                if (newForm.ShowDialog() == DialogResult.OK)
                {
                   
                    this.Close();
                }
            }
            else
            {
                updateFlag = false;
                MessageBox.Show("There is nothing to be updated.\n");
            }
        }

        // Write matched items in Excel
        public void WriteInExcel(bool updateFlag)
        {
            if (updateFlag == true)
            {
                Excel.Application myexcelApplication = new Excel.Application();
                if (myexcelApplication != null)
                {
                    Excel.Workbook myexcelWorkbook = myexcelApplication.Workbooks.Add();
                    Excel.Worksheet myexcelWorksheet = (Excel.Worksheet)myexcelWorkbook.Sheets.Add();

                    const int ITEM_NUMBER = 1;
                    const int ITEM_UPC = 2;
                    const int ITEM_DESC = 3;
                    const int ITEM_PK = 4;
                    const int ITEM_LANDED_COST = 5;
                    const int ITEM_TGP_SRP = 6;
                    const int ITEM_PROGRAM_SRP = 7;
                    const int ITEM_PRICE_DIFF = 8;

                    float margin = (100 - receivedMarginData) / 100f;

                    myexcelWorksheet.Cells[1, ITEM_NUMBER] = "Item Number";
                    myexcelWorksheet.Cells[1, ITEM_UPC] = "Item UPC";
                    myexcelWorksheet.Cells[1, ITEM_DESC] = "Description";
                    myexcelWorksheet.Cells[1, ITEM_PK] = "PK";
                    myexcelWorksheet.Cells[1, ITEM_LANDED_COST] = "Landed Cost";
                    myexcelWorksheet.Cells[1, ITEM_TGP_SRP] = "TGP SRP";
                    myexcelWorksheet.Cells[1, ITEM_PROGRAM_SRP] = "PROGRAM SRP";
                    myexcelWorksheet.Cells[1, ITEM_PRICE_DIFF] = "PRICE DIFF";

                    int row = 2;

                    for (int i = 0; i < updateItemList.Count(); row++, i++)
                    {
                        myexcelWorksheet.Cells[row, ITEM_NUMBER] = updateItemList[i].itemNum;
                        myexcelWorksheet.Cells[row, ITEM_UPC] = updateItemList[i].itemUPC;
                        myexcelWorksheet.Cells[row, ITEM_DESC] = updateItemList[i].itemDesc;
                        myexcelWorksheet.Cells[row, ITEM_PK] = updateItemList[i].itemPk;
                        myexcelWorksheet.Cells[row, ITEM_TGP_SRP] = updateItemList[i].tgp_srp;
                        myexcelWorksheet.Cells[row, ITEM_LANDED_COST] = updateItemList[i].landed_cost;
                        myexcelWorksheet.Cells[row, ITEM_PRICE_DIFF] = updateItemList[i].priceDiff;

                        double landedCost = Convert.ToDouble(updateItemList[i].landed_cost);
                        double itemPk = Convert.ToDouble(updateItemList[i].itemPk);
                        double marginPrice = (landedCost / itemPk) / margin;

                        myexcelWorksheet.Cells[row, ITEM_PROGRAM_SRP] = marginPrice.ToString("0.##");

                        pBar1.PerformStep();
                    }

                    pBar1.Value = pBar1.Maximum;
                    string date = DateTime.Now.ToString("yyyy-MM-dd__hh-mm-ss");
                    tgpExcelFile = tgpExcelFile + date + ".xlsx";

                    myexcelApplication.ActiveWorkbook.SaveAs(tgpExcelFile, Excel.XlFileFormat.xlWorkbookDefault);

                    myexcelWorkbook.Close();
                    myexcelApplication.Quit();

                    MessageBox.Show("An Excel file created.");
                }
            }
            
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            pBar1.Visible = true;
            pBar1.Minimum = 0;
            pBar1.Maximum = databaseItems.Count();
            pBar1.Value = 20;
            pBar1.Step = 1;

            updateFlag = true;

            ReadTGPItems();
            AddToList();
            Compare();
            WriteInExcel(updateFlag);
        }
    }
}
