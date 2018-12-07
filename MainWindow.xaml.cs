using MarketingDataProcessing.Models;
using MarketingDataProcessing.Utilities;
using MarketingDataProcessing.ViewModels;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
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

namespace MarketingDataProcessing
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public string Title = "Update file to database...";
        string Url = @"C:\Users\VuLin\Desktop\DataEditer\DataEditer.xlsx";
        public static string RootDir = @"C:\Users\VuLin\Desktop\DataEditer\WorkSpace\";
        public MainWindow()
        {
            InitializeComponent();
            DataContext = new SearchViewModel();
            var execute = ExcelDataAccess.Execute(Url);
            execute.StartRowInExcel = 2;
            execute.HeaderPosition = 1;
            execute.ExcelToSqlFullColumn<Synthesis>();

        }

        private void LoadFromAExtendFile_Click(object sender, RoutedEventArgs e)
        {

            listInformation.Items.Add("Starting updating data to database..");
            listInformation.Items.Refresh();
            Task<bool> task1 = Task<bool>.Run(() =>
            {
                this.Dispatcher.Invoke((Action)(() =>
                {
                    ShowProcessingSign();
                }));
                try
                {
                    var execute = ExcelDataAccess.Execute(Url);
                    execute.StartRowInExcel = 2;
                    execute.Open<Searching>();
                    return true;
                }
                catch (Exception)
                {
                    return false;
                }
            }).ContinueWith<bool>((theFirstTask) =>
            {
                if (theFirstTask.Result)
                {
                    this.Dispatcher.Invoke((Action)(() =>
                    {
                        HiddenProcessingSign();
                        listInformation.Items.Add("Finish updating data to database...");
                        listInformation.Items.Refresh();
                    }));

                    return true;
                }
                else
                {
                    return false;
                }
            });
            Thread.Sleep(200);

        }
        public void ShowProcessingSign()
        {
            blurGrid.Visibility = Visibility.Visible;
            progressBar.Visibility = Visibility.Visible;
            waitMessage.Visibility = Visibility.Visible;

        }
        public void HiddenProcessingSign()
        {
            blurGrid.Visibility = Visibility.Hidden;
            progressBar.Visibility = Visibility.Hidden;
            waitMessage.Visibility = Visibility.Hidden;
            SearchingResultGrid.Visibility = Visibility.Visible;
        }
    }
}
