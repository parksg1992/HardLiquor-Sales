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
    public static class Log2FileClass
    {
        public static string filePath_temp = System.Reflection.Assembly.GetExecutingAssembly().Location;

        static string year_log = DateTime.Now.ToString("yyyy");
        static string month_log = DateTime.Now.ToString("MM");
        static string date_log = DateTime.Now.ToString("dd");
        static string hour_log = DateTime.Now.ToString("hh");
        static string min_log = DateTime.Now.ToString("mm");
        static string sec_log = DateTime.Now.ToString("ss");
        static string time_log = hour_log + ":" + min_log + ":" + sec_log + "   ";

        static string logFilePath = System.IO.Path.GetDirectoryName(filePath_temp) + "\\Log" + "\\" + year_log + "\\" + month_log + "\\" + date_log + "\\";

        public static void Log2File(string fileName, string content)
        {           
            // If directory does not exist, create it
            if (!Directory.Exists(logFilePath))
            {
                Directory.CreateDirectory(logFilePath);
            }

            System.IO.File.AppendAllText(logFilePath + fileName, time_log + content);
        }
    }
}
