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

namespace KacinProjekat
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public static Excel.Application xlAppRead;
        public static Excel.Workbook xlWorkBookRead;
        public static Excel.Worksheet xlWorkSheetRead;
        public static Excel.Range rangeRead;
        public static int columsRead;
        public static int rowsRead;

        public static Excel.Workbook xlWorkBookWrite;
        public static Excel.Worksheet xlWorkSheetWrite;
        public static Excel.Range rangeWrite;

        public static string path;
        public static string savePath = ConfigurationManager.AppSettings["savingPath"] + DateTime.Now.ToString("dd.MM.yyyy HH,mm All") + ".xlsx";

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

                xlAppRead = new Excel.Application();
                xlWorkBookRead = xlAppRead.Workbooks.Open(path, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlWorkSheetRead = (Excel.Worksheet)xlWorkBookRead.Worksheets.get_Item(1);
                rangeRead = xlWorkSheetRead.UsedRange;
                columsRead = 1;

                xlWorkBookWrite = xlAppRead.Workbooks.Add("");
                xlWorkSheetWrite = (Excel.Worksheet)xlWorkBookWrite.ActiveSheet;

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
            xlWorkBookWrite.SaveAs(savePath, Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, false, false, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            CloseWindowAndFiles();
        }

        private void ButtonFuction(string buttonName)
        {
            if (columsRead > rangeRead.Columns.Count)
            {
                
                xlWorkSheetWrite.Cells[columsRead, 1] = ((string)(rangeRead.Cells[columsRead, rowsRead] as Excel.Range).Value2).TrimEnd().TrimStart();
                xlWorkSheetWrite.Cells[columsRead, 2] = buttonName;

                SeparateExcel separateExcel = new SeparateExcel();

                separateExcel.ShowDialog();
                if (separateExcel.isClosed)
                {
                   xlWorkBookWrite.SaveAs(savePath, Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, false, false, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);                   
                   CloseWindowAndFiles();
                   separateExcel.GenerateFiles(savePath);
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

        private void Button_Click_RollBack(object sender, RoutedEventArgs e)
        {
            columsRead--;
            webBrowser.Navigate(((string)(rangeRead.Cells[columsRead, rowsRead] as Excel.Range).Value2).TrimEnd().TrimStart());
        }

        public void CloseWindowAndFiles()
        {
            xlWorkBookWrite.Close();
            xlWorkBookRead.Close(true, null, null);
            xlAppRead.Quit();
            Marshal.ReleaseComObject(xlWorkSheetWrite);
            Marshal.ReleaseComObject(xlWorkSheetWrite);
            Marshal.ReleaseComObject(xlWorkSheetRead);
            Marshal.ReleaseComObject(xlWorkBookRead);
            Marshal.ReleaseComObject(xlAppRead);

            this.Close();
        }
    }
}
