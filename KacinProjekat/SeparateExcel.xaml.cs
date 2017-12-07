using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace KacinProjekat
{
    /// <summary>
    /// Interaction logic for SeparateExcel.xaml
    /// </summary>
    public partial class SeparateExcel : Window
    {
        public Boolean isClosed { get; set; }
        public SeparateExcel()
        {
            isClosed = false;
            InitializeComponent();          
        }

        private void Button_Click_No(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
        private void Button_Click_Yes(object sender, RoutedEventArgs e)
        {
            this.isClosed = true;
            this.Close();
        }
        public void GenerateFiles(string readPath)
        {
            Excel.Application xlApp = new Excel.Application(); ;
            Excel.Workbook xlWorkBookRead = xlApp.Workbooks.Open(readPath, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            Excel.Worksheet xlWorkSheetRead = (Excel.Worksheet)xlWorkBookRead.Worksheets[1];
            Excel.Range rangeRead = xlWorkSheetRead.UsedRange;

            List<String> categoryList = new List<string>();
            string category;
            for (int i = 1; i <= rangeRead.Columns.Count; i++)
            {
                category = (rangeRead.Cells[i, 2] as Excel.Range).Value2;
                if (!categoryList.Contains(category))
                {
                    Excel.Worksheet xlWorkSheetWrite = (Excel.Worksheet)xlApp.Worksheets.Add(Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                    xlWorkSheetWrite.Name = category + DateTime.Now.ToString(" dd.MM.yyyy hh.mm.ss");
                    Excel.Range rangeWrite = xlWorkSheetWrite.UsedRange;
                    int counter = 0;
                    for (int j = i; j <= rangeRead.Columns.Count; j++)
                    {
                        if ((rangeRead.Cells[j, 2] as Excel.Range).Value2 == category)
                        {
                            counter++;
                            (rangeWrite.Cells[counter, 1] as Excel.Range).Value2 = (rangeRead.Cells[j, 1] as Excel.Range).Value2;
                        }

                    }

                    categoryList.Add(category);
                    ((Excel.Worksheet)xlApp.ActiveWorkbook.Sheets[1]).Activate();
                    xlApp.ActiveWorkbook.Save();
                }

            }

            xlApp.Quit();         
        }
    }
}
