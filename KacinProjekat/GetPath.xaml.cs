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
using System.Windows.Shapes;
using System.IO;

namespace KacinProjekat
{
    /// <summary>
    /// Interaction logic for GetPath.xaml
    /// </summary>
    public partial class GetPath : Window
    {
        public bool closed { get; set; }
        public string path { get; set; }
        public int row { get; set; }
        public GetPath()
        {
            closed = true;
            InitializeComponent();
        }

        private void Button_Click_Continue(object sender, RoutedEventArgs e)
        {
            int rowHelper;
            if (File.Exists(TextBoxUrl.Text))
            {
                if (TextBoxUrl.Text.EndsWith(".xlsx"))
                {
                    if (int.TryParse(TextBoxRow.Text, out rowHelper))
                    {
                        closed = false;
                        row = rowHelper;
                        path = TextBoxUrl.Text;
                        this.Close();
                    }
                    else
                    {
                        MessageBox.Show("Row number that has been inserted is not a number");
                    }
                }
                else
                {
                    MessageBox.Show("Path that you inserted is not path to excel file");
                }
            }
            else
            {
                MessageBox.Show("Path that has been inserted doesen't exist");
            }
        }
        private void Button_Click_Quit(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void Button_Click_Browse(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

            dlg.DefaultExt = ".xlsx";

            Nullable<bool> result = dlg.ShowDialog();

            if (result == true)
            {
                string filename = dlg.FileName;
                TextBoxUrl.Text = filename;
            }
        }
    }
}
