using System;
using System.Collections.Generic;
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
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Configuration;
using System.Reflection;

namespace KacinProjekat
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public static Excel.Application xlApp;
        public static Excel.Workbook xlWorkBookRead;
        public static Excel.Worksheet xlWorkSheetRead;
        public static Excel.Range rangeRead;
        public static int columsRead;
        public static int rowsRead;

        public static Excel.Worksheet xlWorkSheetWrite;

        public static string path;

        public MainWindow()
        {
            GetPath getPath = new GetPath();
            getPath.ShowDialog();

            if (getPath.closed)
            {
                this.Close();
            }
            else
            {
                path = getPath.path;
                rowsRead = getPath.row;
            
                xlApp = new Excel.Application();
                xlWorkBookRead = xlApp.Workbooks.Open(path, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                xlWorkSheetRead = (Excel.Worksheet)xlWorkBookRead.Worksheets.get_Item(xlApp.Sheets.Count);
                rangeRead = xlWorkSheetRead.UsedRange;
                columsRead = 1 ;

                int pera = rangeRead.Columns.Count;

                xlWorkSheetWrite = (Excel.Worksheet)xlApp.Worksheets.Add(Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                xlWorkSheetWrite.Name = "All"+ DateTime.Now.ToString(" dd.MM.yyyy hh.mm.ss");

                InitializeComponent();

                webBrowser.Navigate(((string)(rangeRead.Cells[columsRead, rowsRead] as Excel.Range).Value2).TrimEnd().TrimStart());
            }
        }
        private void Button_Click_Forum(object sender, RoutedEventArgs e)
        {
            ButtonFuction("Forum");    
        }

        private void Button_Click_Blog(object sender, RoutedEventArgs e)
        {
            ButtonFuction("Blog");
        }

        private void Button_Click_Close(object sender, RoutedEventArgs e)
        {
            xlApp.ActiveWorkbook.Save();
            CloseWindowAndFiles();
        }

        private void Button_Click_RollBack(object sender, RoutedEventArgs e)
        {
            columsRead--;
            webBrowser.Navigate(((string)(rangeRead.Cells[columsRead, rowsRead] as Excel.Range).Value2).TrimEnd().TrimStart());
        }

        private void ButtonFuction(string buttonName)
        {
            if (columsRead >= rangeRead.Rows.Count)
            {
                
                xlWorkSheetWrite.Cells[columsRead, 1] = ((string)(rangeRead.Cells[columsRead, rowsRead] as Excel.Range).Value2).TrimEnd().TrimStart();
                xlWorkSheetWrite.Cells[columsRead, 2] = buttonName;

                SeparateExcel separateExcel = new SeparateExcel();

                separateExcel.ShowDialog();
                if (separateExcel.isClosed)
                {
                    xlApp.ActiveWorkbook.Save();
                    CloseWindowAndFiles();
                    separateExcel.GenerateFiles(path);
                }
                
            }
            else
            {
                xlWorkSheetWrite.Cells[columsRead, 1] = ((string)(rangeRead.Cells[columsRead, rowsRead] as Excel.Range).Value2).TrimEnd().TrimStart();
                xlWorkSheetWrite.Cells[columsRead, 2] = buttonName;
                columsRead++;
                webBrowser.Navigate(((string)(rangeRead.Cells[columsRead, rowsRead] as Excel.Range).Value2).TrimEnd().TrimStart());
            }

        }

        public void CloseWindowAndFiles()
        {
            xlWorkBookRead.Close(true, null, null);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlWorkSheetWrite);
            Marshal.ReleaseComObject(xlWorkSheetWrite);
            Marshal.ReleaseComObject(xlWorkSheetRead);
            Marshal.ReleaseComObject(xlWorkBookRead);
            Marshal.ReleaseComObject(xlApp);

            this.Close();
        }
    }
}
