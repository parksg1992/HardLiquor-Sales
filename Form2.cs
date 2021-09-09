using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace HardLiquor_Sales
{
    public partial class Form2 : Form
    {
        public struct ItemInfo
        {
            public string productNumber;
            public string productName;
            public string productPriceChange;
        }

        public static List<ItemInfo> instoreItemInfoList;
        public static List<ItemInfo> newPriceItemInfoList;
        ItemInfo[] instoreItemInfo;
        ItemInfo[] newPriceItemInfo;

        int instoreFileCnt = 0;
        int newPriceFileCnt = 0;

        public const int PRODUCT_NUMBER = 1;
        public const int PRODUCT_NAME = 4;

        OpenFileDialog InstoreOFD = new OpenFileDialog();
        OpenFileDialog NewPriceOFD = new OpenFileDialog();
        
        Excel.Application application;
        Excel.Workbook workbook;
        Excel.Worksheet worksheet;
        object[,] itemNumRawData;
        object[,] itemNameRawData;

        public Form2()
        {
            InitializeComponent();
        }

        public static string instoreFilePath = "";
        public static string newPriceFilePath = "";

        public string SetFilePath(bool isInstoreFile)
        {
            if (isInstoreFile == true)
            {
                return InstoreOFD.FileName;
            }
            else
            {
                return NewPriceOFD.FileName;
            }

        }

        public string GetFilePath(bool isInstoreFile)
        {
            if (isInstoreFile == true)
            {
                return instoreFilePath;
            }
            else
            {
                return newPriceFilePath;
            }
        }

        public int CheckOnSaleFile(ItemInfo instoreItemInfo_)
        {
            for (int i = 0; i < newPriceItemInfoList.Count; i++)
            {
                if (newPriceItemInfoList[i].productNumber.Equals(instoreItemInfo_.productNumber))
                {
                    return i;
                }
            }
            return 0;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (InstoreOFD.ShowDialog() == DialogResult.OK)
            {
                textBox1.Clear();
                textBox1.Text = InstoreOFD.FileName;
                instoreFilePath = SetFilePath(true);
            }

            if (instoreFilePath != "")
            {
                application = new Excel.Application();
                workbook = application.Workbooks.Open(Filename: instoreFilePath);
                worksheet = workbook.Worksheets.get_Item(1);
                application.Visible = false;

                instoreFileCnt = worksheet.UsedRange.Rows.Count;

                // Item Number
                Excel.Range itemNumStartRange = worksheet.Cells[2, 4];
                Excel.Range itemNumEndRange = worksheet.Cells[instoreFileCnt, 4];
                Excel.Range itemNumRange = worksheet.get_Range(itemNumStartRange, itemNumEndRange);
                itemNumRawData = itemNumRange.Value;

                // Item Name
                Excel.Range itemNameStartRange = worksheet.Cells[2, 5];
                Excel.Range itemNameEndRange = worksheet.Cells[instoreFileCnt, 5];
                Excel.Range itemNameRange = worksheet.get_Range(itemNameStartRange, itemNameEndRange);
                itemNameRawData = itemNameRange.Value;

                instoreItemInfo = new ItemInfo[instoreFileCnt];
                instoreItemInfoList = new List<ItemInfo>();
                int idx = 0;

                for (int i = 1; i <= itemNumRawData.GetLength(0); i++)
                {
                    for (int j = 1; j <= itemNumRawData.GetLength(1); j++)
                    {
                        if (itemNumRawData[i, j] == null || itemNameRawData[i, j] == null)
                            continue;

                        instoreItemInfo[idx].productNumber = (itemNumRawData[i, j].ToString());
                        instoreItemInfo[idx].productName = (itemNameRawData[i, j].ToString());

                        break;
                    }

                    instoreItemInfoList.Add(instoreItemInfo[idx]);
                    idx++;
                }

                MessageBox.Show("Upload Done", "Message Box");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (NewPriceOFD.ShowDialog() == DialogResult.OK)
            {
                textBox2.Clear();
                textBox2.Text = NewPriceOFD.FileName;
                newPriceFilePath = SetFilePath(false);
            }

            if (newPriceFilePath != "")
            {
                Excel.Application application_sale = new Excel.Application();
                Excel.Workbook workbook_sale = application_sale.Workbooks.Open(Filename: newPriceFilePath);
                Excel.Worksheet worksheet_sale = workbook_sale.Worksheets.get_Item(1);
                application_sale.Visible = false;

                newPriceFileCnt = worksheet_sale.UsedRange.Rows.Count;

                // Item Number
                Excel.Range itemNumStartRange = worksheet_sale.Cells[6, 3];
                Excel.Range itemNumEndRange = worksheet_sale.Cells[newPriceFileCnt, 3];
                Excel.Range itemNumRange = worksheet_sale.get_Range(itemNumStartRange, itemNumEndRange);
                object[,] itemNumRawData_sale = itemNumRange.Value;

                // Item Name
                Excel.Range itemNameStartRange = worksheet_sale.Cells[6, 4];
                Excel.Range itemNameEndRange = worksheet_sale.Cells[newPriceFileCnt, 4];
                Excel.Range itemNameRange = worksheet_sale.get_Range(itemNameStartRange, itemNameEndRange);
                object[,] itemNameRawData_sale = itemNameRange.Value;

                // Item Price Change
                Excel.Range itemPriceStartRange = worksheet_sale.Cells[6, 19];
                Excel.Range itemPriceEndRange = worksheet_sale.Cells[newPriceFileCnt, 19];
                Excel.Range itemPriceRange = worksheet_sale.get_Range(itemPriceStartRange, itemPriceEndRange);
                object[,] itemPriceRawData_sale = itemPriceRange.Value;

                newPriceItemInfo = new ItemInfo[newPriceFileCnt];
                newPriceItemInfoList = new List<ItemInfo>();
                int idx = 0;
                bool addFlag = false;

                for (int i = 1; i <= itemNumRawData_sale.GetLength(0); i++)
                {
                    for (int j = 1; j <= itemNumRawData_sale.GetLength(1); j++)
                    {
                        if (itemNumRawData_sale[i, j] == null || itemNameRawData_sale[i, j] == null)
                            continue;

                        newPriceItemInfo[idx].productNumber = (itemNumRawData_sale[i, j].ToString());
                        newPriceItemInfo[idx].productName = (itemNameRawData_sale[i, j].ToString());
                        if(itemPriceRawData_sale[i, j] != null)
                        {
                            newPriceItemInfo[idx].productPriceChange = (itemPriceRawData_sale[i, j].ToString());
                        }

                        addFlag = true;

                        break;
                    }
                    if (addFlag == true)
                    {
                        newPriceItemInfoList.Add(newPriceItemInfo[idx]);
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

        private void button3_Click(object sender, EventArgs e)
        {
            int idx = 0;

            for (int i = 1; i <= itemNumRawData.GetLength(0); i++)
            {
                for (int j = 1; j <= itemNumRawData.GetLength(1); j++)
                {
                    if (itemNumRawData[i, j] == null || itemNameRawData[i, j] == null)
                        continue;

                    worksheet.Cells[i + 1, 4].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                    worksheet.Cells[i + 1, 5].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);

                    int index = CheckOnSaleFile(instoreItemInfo[idx]);
                    if (index != 0)
                    {
                        instoreItemInfo[i].productPriceChange = newPriceItemInfoList[index].productPriceChange;

                        // Decrease
                        if ((String.IsNullOrEmpty(instoreItemInfo[i].productPriceChange) == true) || instoreItemInfo[i].productPriceChange.Equals("0"))
                        {
                            // Do Nothing
                        }
                        else if(instoreItemInfo[i].productPriceChange.StartsWith("-"))
                        {
                            worksheet.Cells[i + 1, 4].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                            worksheet.Cells[i + 1, 5].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                            worksheet.Cells[i + 1, 21].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);

                            worksheet.Cells[i + 1, 21] = instoreItemInfo[i].productPriceChange;
                        }
                        // Increase
                        else
                        {
                            worksheet.Cells[i + 1, 4].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            worksheet.Cells[i + 1, 5].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            worksheet.Cells[i + 1, 21].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

                            worksheet.Cells[i + 1, 21] = instoreItemInfo[i].productPriceChange;
                        }
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
}
