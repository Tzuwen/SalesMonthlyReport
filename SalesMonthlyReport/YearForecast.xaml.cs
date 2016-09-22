using System;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using SalesMonthlyReport.AppCode.BLL;
using SalesMonthlyReport.AppCode.BEL;

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
                foreach (DataRowView rowView in dgCustomerDatai1.Items)
                {
                    if (rowView != null)
                    {
                        DataRow row = rowView.Row;
                        objBEL.CustomerId = row.ItemArray[0].ToString();
                        objBEL.Forecast = Convert.ToDecimal(row.ItemArray[2].ToString());
                        objBEL.Year = cbYear.SelectedValue.ToString();
                        YearForecastBLL.insertCustomer(objBEL);
                    }
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
