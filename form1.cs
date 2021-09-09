using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace HardLiquor_Sales
{
    public partial class Form1 : Form
    {
        public struct ItemInfo
        {
            public string productNumber;
            public string productName;
        }
              
        public static List<ItemInfo> orderItemInfoList;
        public static List<ItemInfo> saleItemInfoList;
        ItemInfo[] orderItemInfo;
        ItemInfo[] saleItemInfo;

        int orderFileCnt = 0;
        int saleFileCnt = 0;

        public const int PRODUCT_NUMBER = 1;
        public const int PRODUCT_NAME = 4;

        OpenFileDialog OrderOFD = new OpenFileDialog();
        OpenFileDialog SaleOFD = new OpenFileDialog();


        Excel.Application application;
        Excel.Workbook workbook;
        Excel.Worksheet worksheet;
        object[,] itemNumRawData;
        object[,] itemNameRawData;

        static string year = DateTime.Now.ToString("yyyy");
        static string month = DateTime.Now.ToString("MM");
        static string date = DateTime.Now.ToString("dd");

        public Form1()
        {
            InitializeComponent();
        }

        public static string OrderFilePath = "";
        public static string SaleFilePath = "";

        public string SetFilePath(bool isOrderFile)
        {
            if (isOrderFile == true)
            {
                return OrderOFD.FileName;
            }
            else
            {
                return SaleOFD.FileName;
            }

        }

        public string GetFilePath(bool isOrderFile)
        {
            if (isOrderFile == true)
            {
                return OrderFilePath;
            }
            else
            {
                return SaleFilePath;
            }
        }

        public bool CheckOnSaleFile(ItemInfo orderItemInfo_)
        {
            for(int i = 0; i < saleItemInfoList.Count; i++)
            {
                if(saleItemInfoList[i].productNumber.Equals(orderItemInfo_.productNumber))
                {
                    return true;
                }
            }
            return false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (OrderOFD.ShowDialog() == DialogResult.OK)
            {
                textBox1.Clear();
                textBox1.Text = OrderOFD.FileName;
                OrderFilePath = SetFilePath(true);
            }
            
            string sourceFile = OrderFilePath;
            string destinationFile = sourceFile + "_"+ year + "_" + month + "_" + date + ".xlsx";

            try
            {
                File.Copy(sourceFile, destinationFile, true);
            }
            catch (IOException iox)
            {
                Console.WriteLine(iox.Message);
            }

            if (destinationFile != "")
            {
                application = new Excel.Application();
                workbook = application.Workbooks.Open(Filename: destinationFile);
                worksheet = workbook.Worksheets.get_Item(1);
                application.Visible = false;

                orderFileCnt = worksheet.UsedRange.Rows.Count;

                // Item Number
                Excel.Range itemNumStartRange = worksheet.Cells[2, 1];
                Excel.Range itemNumEndRange = worksheet.Cells[orderFileCnt, 1];
                Excel.Range itemNumRange = worksheet.get_Range(itemNumStartRange, itemNumEndRange);
                itemNumRawData = itemNumRange.Value;

                // Item Name
                Excel.Range itemNameStartRange = worksheet.Cells[2, 3];
                Excel.Range itemNameEndRange = worksheet.Cells[orderFileCnt, 3];
                Excel.Range itemNameRange = worksheet.get_Range(itemNameStartRange, itemNameEndRange);
                itemNameRawData = itemNameRange.Value;

                orderItemInfo = new ItemInfo[orderFileCnt];
                orderItemInfoList = new List<ItemInfo>();
                int idx = 0;

                for (int i = 1; i <= itemNumRawData.GetLength(0); i++)
                {
                    for (int j = 1; j <= itemNumRawData.GetLength(1); j++)
                    {
                        if (itemNumRawData[i, j] == null || itemNameRawData[i, j] == null)
                            continue;

                        orderItemInfo[idx].productNumber = (itemNumRawData[i, j].ToString());
                        orderItemInfo[idx].productName = (itemNameRawData[i, j].ToString());

                        break;
                    }

                    orderItemInfoList.Add(orderItemInfo[idx]);
                    idx++;
                }

                MessageBox.Show("Upload Done", "Message Box");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (SaleOFD.ShowDialog() == DialogResult.OK)
            {
                textBox2.Clear();
                textBox2.Text = SaleOFD.FileName;
                SaleFilePath = SetFilePath(false);
            }

            if (SaleFilePath != "")
            {
                Excel.Application application_sale = new Excel.Application();
                Excel.Workbook workbook_sale = application_sale.Workbooks.Open(Filename: SaleFilePath);
                Excel.Worksheet worksheet_sale = workbook_sale.Worksheets.get_Item(1);
                application_sale.Visible = false;

                saleFileCnt = worksheet_sale.UsedRange.Rows.Count;

                // Item Number
                Excel.Range itemNumStartRange = worksheet_sale.Cells[3, 2];
                Excel.Range itemNumEndRange = worksheet_sale.Cells[saleFileCnt, 2];
                Excel.Range itemNumRange = worksheet_sale.get_Range(itemNumStartRange, itemNumEndRange);
                object[,] itemNumRawData_sale = itemNumRange.Value;

                // Item Name
                Excel.Range itemNameStartRange = worksheet_sale.Cells[3, 3];
                Excel.Range itemNameEndRange = worksheet_sale.Cells[saleFileCnt, 3];
                Excel.Range itemNameRange = worksheet_sale.get_Range(itemNameStartRange, itemNameEndRange);
                object[,] itemNameRawData_sale = itemNameRange.Value;

                saleItemInfo = new ItemInfo[saleFileCnt];
                saleItemInfoList = new List<ItemInfo>();
                int idx = 0;
                bool addFlag = false;

                for (int i = 1; i <= itemNumRawData_sale.GetLength(0); i++)
                {
                    for (int j = 1; j <= itemNumRawData_sale.GetLength(1); j++)
                    {
                        if (itemNumRawData_sale[i, j] == null || itemNameRawData_sale[i, j] == null)
                            continue;

                        saleItemInfo[idx].productNumber = (itemNumRawData_sale[i, j].ToString());
                        saleItemInfo[idx].productName = (itemNameRawData_sale[i, j].ToString());
                        addFlag = true;

                        break;
                    }
                    if (addFlag == true)
                    {
                        saleItemInfoList.Add(saleItemInfo[idx]);
                        idx++;
                        addFlag = false;
                    }
                }

                MessageBox.Show("Upload done.", "Message Box");

                DeleteObject(worksheet_sale);

                workbook_sale.Close();
                DeleteObject(workbook_sale);

                application_sale.Quit();
                DeleteObject(application_sale);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (textBox3.Text == "")
            {
                MessageBox.Show("Please enter Month", "Message Box");
            }
            else
            {

                int idx = 0;

                for (int i = 1; i <= itemNumRawData.GetLength(0); i++)
                {
                    for (int j = 1; j <= itemNumRawData.GetLength(1); j++)
                    {
                        if (itemNumRawData[i, j] == null || itemNameRawData[i, j] == null)
                            continue;

                        /*if (worksheet.Cells[i + 1, 2] != null)
                        {
                            worksheet.Cells[i + 1, 2] = null;
                        }*/

                        if (CheckOnSaleFile(orderItemInfo[idx]) == true)
                        {
                            string month = textBox3.Text;
                            string str = worksheet.Cells[i + 1, 3].Value.ToString();
                            string saleMonth = "___" + month;

                            worksheet.Cells[i + 1, 3] = str + "  ___" + month;

                            // Colour Change
                            Excel.Range changeColour = worksheet.Cells[i + 1, 3];
                            changeColour.Value = string.Format("{0},{1}", str, saleMonth);
                            changeColour.Characters[str.Length + 1, str.Length + 1 + saleMonth.Length].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                        }

                        break;
                    }
                    idx++;
                }

                workbook.Save();

                MessageBox.Show("Done", "Message Box");

                DeleteObject(worksheet);

                workbook.Close();
                DeleteObject(workbook);

                application.Quit();
                DeleteObject(application);
            }
        }

        static void DeleteObject(object obj)
        {
            try
            {
                if (obj != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                    obj = null;
                }
            }
            catch (Exception e)
            {
                obj = null;
                throw e;
            }
            finally
            {
                GC.Collect();
            }
        }

        
    }

}
