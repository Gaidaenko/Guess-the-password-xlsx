using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Guess_the_password_xlsx
{
    public partial class Form1 : Form
    {
        public static string tryPassword;

        public static string file;
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.Filter = "Формат xlsx(*.xlsx)|*.xlsx|Все файлы(*.*)|*.*";
            openFile.Title = "Выберете файл";

            if (openFile.ShowDialog() == DialogResult.OK)
            {
               file = openFile.FileName;
            }

            Excel.Application xlsApp = new Excel.Application();
            Workbook ObjWorkBook = xlsApp.Workbooks.Open(file, 0, false, 5, tryPassword, "", false, XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            Worksheet ObjWorkSheet;
        }

        private void label1_Click(object sender, EventArgs e)
        {
           
        }
    }
}
