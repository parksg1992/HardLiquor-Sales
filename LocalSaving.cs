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
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using Excel = Microsoft.Office.Interop.Excel;

namespace HardLiquor_Sales
{
    public partial class LocalSaving : Form
    {
        public struct ItemInfo
        {
            public string itemName;
            public string itemOrderNumber;
            public string itemUPC;
            public string itemCase;
            public string itemSave;
            public int itemQuantity;
        }
        List<ItemInfo> databaseItems = new List<ItemInfo>();
        List<ItemInfo> orderItems = new List<ItemInfo>();
        List<ItemInfo> creditItems = new List<ItemInfo>();

        public LocalSaving()
        {
            InitializeComponent();
            Read();
        }
        OpenFileDialog localSavingOFD = new OpenFileDialog();

        public string orderDate;
        public string orderDateMonth;
        public string orderDateDate;
        public string orderDateYear;

        public static string filePath_temp = System.Reflection.Assembly.GetExecutingAssembly().Location;
        public string dbFilePath = System.IO.Path.GetDirectoryName(filePath_temp) + "\\Sales_Database.xml";
        public string localSavingExcelFile = System.IO.Path.GetDirectoryName(filePath_temp) + "\\Local Saving_";

        public static bool noMoreFileToAddFlag = false;

        public void Read()
        {
            XmlDocument xml = new XmlDocument();
            xml.Load(dbFilePath);
            XmlNode xmlList_1 = xml.SelectNodes("items")[0];

            foreach (XmlNode xnl in xmlList_1)
            {
                var temp = new ItemInfo();
                temp.itemName = xnl.Attributes["name"].Value;
                temp.itemOrderNumber = xnl.Attributes["order_num"].Value;
                temp.itemUPC = xnl.Attributes["upc"].Value;
                temp.itemCase = xnl.Attributes["case"].Value;
                temp.itemSave = xnl.Attributes["save"].Value;
                temp.itemQuantity = 0;

                databaseItems.Add(temp);
            }
        }


        // Orderfile Upload (PDF Read)
        private void button1_Click(object sender, EventArgs e)
        {
            while (noMoreFileToAddFlag == false)
            {
                if (localSavingOFD.ShowDialog() == DialogResult.OK)
                {
                    textBox1.Clear();
                    textBox1.Text = localSavingOFD.FileName;
                }

                StringBuilder sb = new StringBuilder();

                using (PdfReader reader = new PdfReader(localSavingOFD.FileName))
                {
                    for (int pageNum = 1; pageNum <= reader.NumberOfPages; pageNum++)
                    {
                        ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
                        string text = PdfTextExtractor.GetTextFromPage(reader, pageNum, strategy);
                        text = Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(text)));

                        string[] parse = text.Split('\n');

                        int idx = 0;
                        var temp = new ItemInfo();

                        while (true)
                        {
                            if (parse.Count() == idx)
                            {
                                break;
                            }

                            var str = parse[idx];

                            // Get the order date
                            if (pageNum == 1 && idx == 6)
                            {
                                idx++;
                                orderDate = str;

                                string[] orderDateParse = orderDate.Split(' ');
                                string[] orderDateParse2 = orderDateParse[2].Split('/');
                                orderDateMonth = orderDateParse2[0];
                                orderDateDate = orderDateParse2[1];
                                orderDateYear = orderDateParse2[2];


                                // need to parse orderDate
                                continue;
                            }

                            // Dump data
                            if (pageNum == 1 && idx < 11)
                            {
                                idx++;
                                continue;
                            }

                            if (pageNum != 1 && idx < 5)
                            {
                                idx++;
                                continue;
                            }

                            // Item OrderNumber
                            
                            if(str.All(char.IsDigit) == false)
                            {
                                break;
                            }
                            temp.itemOrderNumber = str;
                            idx++;

                            // Item Name
                            temp.itemName = parse[idx].ToString();
                            idx++;
                            

                            // Item Quantity
                            string tempNumber = parse[idx].ToString();
                            string[] parse2 = tempNumber.Split(' ');
                            string tempItemQuantity = parse2[2].ToString();
                            temp.itemQuantity = Int32.Parse(tempItemQuantity);

                            idx++;

                            orderItems.Add(temp);

                            if (pageNum == reader.NumberOfPages && idx == parse.Count() - 1)
                            {
                                break;
                            }
                        }
                    }
                }


                Find();

                if (noMoreFileToAddFlag == true)
                {
                    WriteInExcel();
                }
            }
        }

        // Find matched items
        public void Find()
        {
            // databaseItems
            // orderItems

            for (int idx = 0; idx < databaseItems.Count(); idx++)
            {
                var searchForOrderNum = databaseItems[idx].itemOrderNumber;

                var matchedList = orderItems.Find(x => x.itemOrderNumber == searchForOrderNum);

                if (matchedList.itemOrderNumber != null)
                {
                    // 이미 CreditItems List에 있는지 확인 후 있으면 Quantity ++
                    var existCheck = creditItems.Find(x => x.itemOrderNumber == matchedList.itemOrderNumber);
                    

                    // 이미 존재한다
                    if (existCheck.itemOrderNumber != null)
                    {
                        int findIndex = creditItems.FindIndex(x => x.itemOrderNumber == matchedList.itemOrderNumber);
                        //int itemQuantity = matchedList.itemQuantity + 1;
                        
                        ItemInfo temp = creditItems[findIndex];
                        temp.itemName = creditItems[findIndex].itemName;
                        temp.itemCase = creditItems[findIndex].itemCase;
                        temp.itemOrderNumber = creditItems[findIndex].itemOrderNumber;
                        temp.itemSave = creditItems[findIndex].itemSave;
                        temp.itemUPC = creditItems[findIndex].itemUPC;
                        temp.itemQuantity = creditItems[findIndex].itemQuantity + 1;

                        creditItems[findIndex] = temp;
                        
                    }
                    else
                    {
                        // List 에 추가
                        matchedList.itemCase = databaseItems[idx].itemCase;
                        matchedList.itemSave = databaseItems[idx].itemSave;
                        matchedList.itemUPC = databaseItems[idx].itemUPC;
                        creditItems.Add(matchedList);
                    }
                    
                }
            }
            
            // 더 있니? 메시지 창
            MoreFileToAdd mfta = new MoreFileToAdd();
            mfta.ShowDialog();

            if(noMoreFileToAddFlag == true)
            {
                this.Close();
            }
        }

        // Write matched items in Excel
        public void WriteInExcel()
        {
            Excel.Application myexcelApplication = new Excel.Application();
            if (myexcelApplication != null)
            {
                Excel.Workbook myexcelWorkbook = myexcelApplication.Workbooks.Add();
                Excel.Worksheet myexcelWorksheet = (Excel.Worksheet)myexcelWorkbook.Sheets.Add();

                const int ITEM_NAME = 1;
                const int ORDER_NUMBER = 2;
                const int CASE = 3;
                const int SAVING = 4;
                const int QUANTITY = 5;
                const int TOTAL = 6;
                const int TOTAL_SAVINGS = 7;


                int creditListIndex = 0;

                myexcelWorksheet.Cells[1, ITEM_NAME] = "Item";
                myexcelWorksheet.Cells[1, ORDER_NUMBER] = "Order Number";
                myexcelWorksheet.Cells[1, CASE] = "Case";
                myexcelWorksheet.Cells[1, SAVING] = "Saving";
                myexcelWorksheet.Cells[1, QUANTITY] = "Order Amount";
                myexcelWorksheet.Cells[1, TOTAL] = "Total";
                myexcelWorksheet.Cells[1, TOTAL_SAVINGS] = "Total Savings";
                
                for (int row = 2; creditListIndex < creditItems.Count(); row++, creditListIndex++)
                {
                    myexcelWorksheet.Cells[row, ITEM_NAME] = creditItems[creditListIndex].itemName;
                    myexcelWorksheet.Cells[row, ORDER_NUMBER] = creditItems[creditListIndex].itemOrderNumber;
                    myexcelWorksheet.Cells[row, CASE] = creditItems[creditListIndex].itemCase;
                    myexcelWorksheet.Cells[row, SAVING] = creditItems[creditListIndex].itemSave;
                    myexcelWorksheet.Cells[row, QUANTITY] = creditItems[creditListIndex].itemQuantity;
                    myexcelWorksheet.Cells[row, TOTAL] = creditItems[creditListIndex].itemQuantity * Int32.Parse(creditItems[creditListIndex].itemCase);
                    myexcelWorksheet.Cells[row, TOTAL_SAVINGS] = (creditItems[creditListIndex].itemQuantity * Int32.Parse(creditItems[creditListIndex].itemCase) 
                        * Convert.ToDouble(creditItems[creditListIndex].itemSave));
                }
                string date = DateTime.Now.ToString("yyyy-MM-dd__hh-mm-ss");
                localSavingExcelFile = localSavingExcelFile + date + ".xlsx";

                myexcelApplication.ActiveWorkbook.SaveAs(localSavingExcelFile, Excel.XlFileFormat.xlWorkbookDefault);

                myexcelWorkbook.Close();
                myexcelApplication.Quit();
            }


        }
    }
}