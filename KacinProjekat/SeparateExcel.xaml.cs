﻿using System;
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
        public void GenerateFiles(string path)
        {
            // TO DO
        }
    }
}
