using System;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Linq;
using System.Diagnostics;
using System.IO;
using System.Threading;

namespace Guess_the_password_xlsx
{
    public partial class Form1 : Form
    {
        public static string tryPassword;                                                                          
        public static string [] textArr;                                                                            
        public static string fileXlsx;
        public static string fileTxt;
        public static int i = 0;
        private delegate void SafeCallDelegate(string text);
        private Thread thread = null;

        public Form1()
        {
            InitializeComponent();                                             
        }

        private void openTxt()
        {
            OpenFileDialog openFileTxt = new OpenFileDialog();                                       //сделать исключение при отсутствии выбора файла txt
            openFileTxt.Filter = "Формат txt(*.txt)|*.txt";                                           // сделать кнопку -  превать подбор пароля
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
        private void WriteTextLabel_4(string textLabel4)
        {
            if (label1.InvokeRequired)
            {
                var lb1 = new SafeCallDelegate(WriteTextLabel_4);
                label4.Invoke(lb1, new object[] {textLabel4});
            }
            else
            {
                label4.Text = tryPassword;
            }
        }
        private void WriteTextLabel_3(string textLabel2)
        {
            
            if (label2.InvokeRequired)
            {
                var lb2 = new SafeCallDelegate(WriteTextLabel_3);
                label2.Invoke(lb2, new object[] {textLabel2});
            }
            else
            {
                label2.Text = i.ToString();
            }
        }

        public void Try()
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

                            WriteTextLabel_4(i.ToString());    
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
                            WriteTextLabel_3(tryPassword);

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
            Thread thread2 = new Thread(Try);           
            thread2.Start();           
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
