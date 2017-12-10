using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
      
        private void Form1_Load(object sender, EventArgs e)
        {

        }
        public int numOfLastRow = 0;
        private void button1_Click(object sender, EventArgs e)
        {
            //button2.Enabled = false;
            Stream stream = null;
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.InitialDirectory = "D:\\";
            dialog.Filter = "Excel files (*.xls, *.xlsx, *.xlsm)|*.xls;*.xlsx;*.xlsm";
        
         
            if (dialog.ShowDialog() == DialogResult.OK)
                try
                {
                    Microsoft.Office.Interop.Excel.Application myExcel;
                    Microsoft.Office.Interop.Excel.Workbook myWorkbook;
                    Microsoft.Office.Interop.Excel.Worksheet myWorksheet;

                    myExcel = new Microsoft.Office.Interop.Excel.Application();
                    myExcel.Workbooks.Open(@dialog.FileName);
                    myWorkbook = myExcel.ActiveWorkbook;
                    myWorksheet = (Excel.Worksheet)myWorkbook.Sheets[1];
                    
                    StreamWriter sw = new StreamWriter(@"D:\new.txt",false,Encoding.GetEncoding("Windows-1251"));

                    string str = "";
                    int count = 0;
                    var lastCell = myWorksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);

                    //записываем название, автора и вопросы
                    for (int i = 0; i < lastCell.Column; i++)
                    {
                        for (int j = 0; j < lastCell.Row; j++)
                        {
                            str = myWorksheet.Cells[j + 1, i + 1].Text.ToString();
                            if ((myWorksheet.Cells[j + 1, 1].Text.ToString() == "") || (myWorksheet.Cells[j + 1, 1].Text.ToString() == "    "))
                                count++;
                            sw.WriteLine(str);
                            if (count == 2)
                            {
                                numOfLastRow = j + 1;
                                break;
                            }
                        }
                        if (count == 2) break;
                    }
                 

                   
                    //записываем вероятности
                    for (int j = numOfLastRow; j < lastCell.Row; j++)
                    {
                        for (int i = 0; i < lastCell.Column; i++)
                        {
                            str = myWorksheet.Cells[j + 1, i + 1].Text.ToString();
                            str = str.Replace(",", ".");
                            if ((myWorksheet.Cells[j + 1, i + 2].Text.ToString() == "") || (myWorksheet.Cells[j + 1, i + 2].Text.ToString() == "    "))
                                str += "\n";
                            else
                                str += ",";
                            sw.Write(str);
                        }
                    }

                    sw.Close();
                    myWorkbook.Close(false);
                    myExcel.Quit();


                }
                catch (Exception ex)
                {
                    label1.Text = "Файл не выбран"; 
                }
                string name = dialog.FileName;
                int position = name.LastIndexOf("\\");
                name = name.Substring(position + 1);
                label1.Text = name;
            MessageBox.Show("Готово");

        }

        private void button2_Click(object sender, EventArgs e)
        {
            //button1.Enabled = false;
            Stream myStream = null;
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = "D:\\";
            openFileDialog1.Filter = "txt files (.txt)|*.txt";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;
            File.Copy(@"D:\new.txt", @"D:\new1.txt");

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    if ((myStream = openFileDialog1.OpenFile()) != null)
                    {
                        using (myStream)
                        {
                            string resultString = "";
                            System.IO.StreamReader rw = new System.IO.StreamReader(myStream);
                            string line;
                            while ((line = rw.ReadLine()) != null)
                            {
                                resultString += line;
                            }

                        }
                    }
                }
                catch (Exception)
                {
                    label2.Text = "Файл не выбран ";
                }
                string name = openFileDialog1.FileName;
                int position = name.LastIndexOf("\\");
                name = name.Substring(position + 1);
                label2.Text = name;
                
                MessageBox.Show("Готово");
                System.IO.File.Move(@"D:\new1.txt", @"D:\new2.mkb");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Excel должен быть сохранен отдельной книгой. При открытии файла, поддерживаются форматы .xls .xlsx .xlsm. При нажатии на кнопку нужно открыть нужный файл, после чего создается новый с именем 'new' в формате .txt. Чтобы преобразовать файл в формат .mkb. Нажмите кнопку справа и выберите файл 'new1'(это копия файла формата .txt) ");
        }
    }
}
