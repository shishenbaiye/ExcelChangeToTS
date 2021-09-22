using System;
using System.Collections.Generic;
using System.Data;
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
using Microsoft.Win32;

namespace WpfApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {   
        private static FileDialog _file;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog();
            dialog.Filter = "Excel文件|*.xls;*.xlsx";
            if (dialog.ShowDialog(this) == false) return;
            var _fileName = dialog.FileName;
            this.FileName.Content = System.IO.Path.GetFileName(_fileName);
            _file = dialog;
            
        }

        private void ProgressBar_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {   
           
        }

        private void textBox_TextChanged(object sender, TextChangedEventArgs e)
        {

        }
        private void ChangeButton_Click(object sender, RoutedEventArgs e)
        {
            ExcelData.ProcessExcel.Excel(_file.FileName);
        }

    }
}
