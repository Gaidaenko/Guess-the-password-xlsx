using System;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Linq;
using System.Diagnostics;
using System.IO;


namespace Guess_the_password_xlsx
{
    public partial class Form1 : Form
    {
        public static string tryPassword;                                                                          
        public static string [] textArr;                                                                            
        public static string fileXlsx;
        public static string fileTxt;
        public static int i = 0;
        public Form1()
        {
            InitializeComponent();
        }

        private void openTxt()
        {
            OpenFileDialog openFileTxt = new OpenFileDialog();
            openFileTxt.Filter = "Формат txt(*.txt)|*.txt";                 
            openFileTxt.Title = "Выберете файл";

            if (openFileTxt.ShowDialog() == DialogResult.OK)
            {
                fileTxt = openFileTxt.FileName;
            }
        }
        private void openXlsx()
        {
            OpenFileDialog openFileXlsx = new OpenFileDialog();
            openFileXlsx.Filter = "Формат xlsx(*.xlsx)|*.xlsx|xls(*.xls)|*.xls";                
            openFileXlsx.Title = "Выберете файл";

            if (openFileXlsx.ShowDialog() == DialogResult.OK)
            {
                fileXlsx = openFileXlsx.FileName;
            }
        }
        private void Try()
        {
            try
            {
                using (StreamReader text = new StreamReader(fileTxt))
                {
                    var result = text.ReadToEnd();
                    result.ToArray();

                    if (String.IsNullOrEmpty(result))
                    {
                        MessageBox.Show("Вероятно, файл с вариантами паролей пуст.");
                        return;
                    }

                    textArr = result.Split();

                Next: for (i = i; i < textArr.Length; i++)
                    {
                        tryPassword = textArr[i];

                        try
                        {
                            Excel.Application xlsApp = new Excel.Application();
                            Workbook ObjWorkBook = xlsApp.Workbooks.Open
                                (Filename: fileXlsx,
                                 UpdateLinks: 0,
                                 ReadOnly: true,
                                 Format: 5,
                                 Password: tryPassword,
                                 WriteResPassword: true,
                                 false,
                                 Origin: XlPlatform.xlWindows,
                                 Delimiter: "",
                                 Editable: true,
                                 Notify: false,
                                 Converter: 0,
                                 AddToMru: true,
                                 Local: false,
                                 CorruptLoad: false);

                            Process.Start(fileXlsx);                                                                                      

                            label4.Text = tryPassword;
                            return;
                        }
                        catch
                        {
                            Process[] List;
                            List = Process.GetProcessesByName("EXCEL");
                            foreach (var process in List)
                            {
                                process.Kill();
                            }

                            i++;
                            label2.Text = i.ToString();

                           if (textArr[i] == null)                                                              
                           {
                                //go to exception                                                                     
                           }

                            goto Next; 
                        }
                    }
                }
            }
            catch(Exception e)
            {
                MessageBox.Show("Отсутствуют подходящие пароли.");
            }
        }            

        private void button1_Click(object sender, EventArgs e)
        {
            openXlsx();          
        }
        private void button2_Click(object sender, EventArgs e)
        {
            openTxt();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            Try();
        }
        private void label1_Click(object sender, EventArgs e)
        {

        }
        private void label2_Click(object sender, EventArgs e)
        {

        }
        private void label3_Click(object sender, EventArgs e)
        {

        }
        private void label4_Click(object sender, EventArgs e)
        {

        }
        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
