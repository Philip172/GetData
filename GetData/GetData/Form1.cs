using Microsoft.Office.Interop.Excel; //Подключаем соответствующую библиотеку
using System;
using System.Windows.Forms;

namespace GetData
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            progressBar1.Value++;

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application(); //Excel
            Workbook xlWB; //рабочая книга откуда будем копировать лист  
            Workbook xlWB2; //рабочая книга куда будем копировать лист
            Worksheet xlSht; //лист Excel            

            progressBar1.Value++;

            //название файла Excel откуда будем копировать лист
            xlWB = xlApp.Workbooks.Open(@"C:\C#\FromHere.xlsx");

            progressBar1.Value++;

            //название файла Excel куда будем копировать лист
            xlWB2 = xlApp.Workbooks.Open(@"C:\C#\ToHere.xlsx");

            progressBar1.Value++;

            //название листа или 1-й лист в книге xlSht = xlWB.Worksheets[1];
            xlSht = xlWB.Worksheets[1];

            progressBar1.Value++;

            //сам процесс копирования листа из одного файла в другой
            xlSht.Copy(After: xlWB2.Worksheets[xlWB2.Worksheets.Count]);

            progressBar1.Value++;

            MessageBox.Show("Лист '" + xlSht.Name.ToString() + "' успешно скопирован", "Поиск", MessageBoxButtons.OK, MessageBoxIcon.Information);

            progressBar1.Value++;

            xlApp.Visible = true; //отображаем Excel

            progressBar1.Value++;

            //xlWB2.Close(true); //закрываем и сохраняем изменения в файле 2            


            //xlApp.Quit(); //закрываем Excel



            progressBar1.Value = progressBar1.Maximum;
        }
    }
}
