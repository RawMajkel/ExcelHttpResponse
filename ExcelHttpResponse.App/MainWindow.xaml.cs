using ExcelHttpResponse.Services;
using Microsoft.Win32;
using System.Windows;

namespace ExcelHttpResponse.App
{
    public partial class MainWindow : Window
    {
        private static ExcelService _excelService;

        public MainWindow()
        {
            InitializeComponent();

            _excelService = new ExcelService();
        }

        private void sendXlsx_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Xlsx files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                var path = openFileDialog.FileName;
                _excelService.GetAndSaveResultsToFile(path);
                MessageBox.Show("Gotowe", "Sukces - zapisano plik");
            }
        }
    }
}
