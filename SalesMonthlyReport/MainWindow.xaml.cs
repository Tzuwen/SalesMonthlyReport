using System;
using System.Windows;

namespace SalesMonthlyReport
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            LoadPage("MonthlyReport");
        }

        private void LoadPage(string page)
        {
            frame1.Source = new Uri(page + ".xaml", UriKind.RelativeOrAbsolute);
        }

        private void MonthlyReport_Click(object sender, RoutedEventArgs e)
        {
            LoadPage("MonthlyReport");
        }

        private void CustomerSetup_Click(object sender, RoutedEventArgs e)
        {
            LoadPage("CustomerSetup");
        }

        private void YearForecast_Click(object sender, RoutedEventArgs e)
        {
            LoadPage("YearForecast");
        }
    }
}
