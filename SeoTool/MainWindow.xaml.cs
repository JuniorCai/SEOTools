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
using PublicFramework;

namespace SeoTool
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.Title = "选择文件";
            openFileDialog.Filter = "Excel文件(*.xls,*.xlsx)|*.xls;*.xlsx";
            openFileDialog.FileName = "选择文件夹.";
            openFileDialog.FilterIndex = 1;
            openFileDialog.ValidateNames = false;
            openFileDialog.CheckFileExists = false;
            openFileDialog.CheckPathExists = true;
            openFileDialog.Multiselect = false; //允许同时选择多个文件 
            bool? result = openFileDialog.ShowDialog();
            if (result != true)
            {
                return;
            }
            else
            {
                string[] files = openFileDialog.FileNames;
                if (files.Length > 0)
                {
                    filePath.Text = files[0];

                    
                }

            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            string fileUrl = filePath.Text;


            ExcelOperator excelOperator = new ExcelOperator(fileUrl, rootWord.Text);
            excelOperator.Open();
            excelOperator.Read();
            excelOperator.ApartOriginalWords();

            excelOperator.Close();
        }
    }
}
