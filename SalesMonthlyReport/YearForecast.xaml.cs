using System;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using SalesMonthlyReport.AppCode.BLL;
using SalesMonthlyReport.AppCode.BEL;
using SalesMonthlyReport.AppCode;

namespace SalesMonthlyReport
{
    /// <summary>
    /// YearForecast.xaml 的互動邏輯
    /// </summary>
    public partial class YearForecast : Page
    {
        public YearForecast()
        {
            InitializeComponent();
            Initialize();
        }

        private void Initialize()
        {
            // Set year dropdown
            int year = DateTime.Now.Year + 1;
            for (int i = 0; i < 3; i++)
                cbYear.Items.Add((year - i).ToString());
            cbYear.SelectedIndex = 0;

            // Set sales dropdown
            DataTable dt = SalesBLL.getSalesAll();
            foreach (DataRow dr in dt.Rows)
                cbSales.Items.Add(dr[0].ToString() + "-" + dr[1].ToString());
            cbSales.SelectedIndex = 0;
            dt = null;
        }

        private void SetCustomerDg()
        {
            try
            {
                // Set customer year forecast datagrid
                string selectedYear = cbYear.SelectedValue.ToString();
                string selectedSales = (cbSales.SelectedValue.ToString().Split('-'))[0].ToString();
                DataTable dt = YearForecastBLL.getYearForecastByYearSales(selectedYear, selectedSales);
                dgCustomerDatai1.ItemsSource = dt.DefaultView;
                dt = null;
            }
            catch
            {

            }
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                YearForecastBEL objBEL = new YearForecastBEL();
                for (int i = 0; i < dgCustomerDatai1.Items.Count; i++)
                {
                    DataGridCell cell0 = FunctionClass.GetCell(dgCustomerDatai1, i, 0);                   
                    DataGridCell cell2 = FunctionClass.GetCell(dgCustomerDatai1, i, 2);
                    TextBlock tb0 = cell0.Content as TextBlock;
                    TextBlock tb1 = cell2.Content as TextBlock;
                    string customerId = tb0.Text.Trim();
                    string forecast = tb1.Text.Trim() == "" ? "0" : tb1.Text.Trim();                    
                    objBEL.CustomerId = tb0.Text.Trim();
                    objBEL.Forecast = Convert.ToDecimal(tb1.Text.Trim());
                    objBEL.Year = cbYear.SelectedValue.ToString();
                    YearForecastBLL.insertCustomer(objBEL);                    
                }
               
                System.Windows.Forms.MessageBox.Show("設定完畢", "Message");
            }
            catch
            {

            }
        }

        private void cbYear_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            SetCustomerDg();
        }

        private void cbSales_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            SetCustomerDg();
        }
    }
}
