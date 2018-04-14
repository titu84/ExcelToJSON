using Microsoft.Win32;
using ReadDataFromExcel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
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
using System.Web.Script.Serialization;

namespace Excel2Json
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        string patch;
        string json;
        private void btnGetExcel_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel|*.xls;*.xlsx;*.xlsm";
            if (openFileDialog.ShowDialog() == true)
                patch = openFileDialog.FileName;
            tbFileName.Text = patch;
            btnCreateJson.IsEnabled = true;
            btnSaveJson.IsEnabled = false;
        }

        private void btnCreateJson_Click(object sender, RoutedEventArgs e)
        {
            btnCreateJson.IsEnabled = false;
            ExcelReader gef = new ExcelReader(patch);
            json = gef.GetJsonFormExcel(1);
            richTextBox.Document.Blocks.Clear();
            richTextBox.Document.Blocks.Add(new Paragraph(new Run(json)));
            richTextBox.IsEnabled = false;            
            btnSaveJson.IsEnabled = true;
        }

        private void btnSaveJson_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "JSON file (*.json)|*.json";
            if (saveFileDialog.ShowDialog() == true)
                File.WriteAllText(saveFileDialog.FileName, json);
            MessageBoxResult result = MessageBox.Show("Zapisano");
            
        }
    }
}
