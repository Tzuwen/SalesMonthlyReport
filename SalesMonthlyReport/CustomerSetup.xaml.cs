using SalesMonthlyReport.AppCode.BEL;
using SalesMonthlyReport.AppCode.BLL;
using System.Data;
using System.Windows;
using System.Windows.Controls;

namespace SalesMonthlyReport
{
    public partial class CustomerSetup : Page
    {
        public CustomerSetup()
        {
            InitializeComponent();
            Initialize();
        }

        private void Initialize()
        {
            btnMoveRight.Content = ">>";
            btnMoveLeft.Content = "<<";            
            SetCurrentSales();
            SetCustomerGrid1();
            SetCustomerGrid2();
        }

        private void SetCurrentSales()
        {
            DataTable dt = SalesBLL.getSalesAll();
            dgCurrentSales.ItemsSource = dt.DefaultView;
            cbSales01.Items.Clear();
            cbSales02.Items.Clear();
            foreach (DataRow dr in dt.Rows)
            {
                cbSales01.Items.Add(dr[0].ToString() + "-" + dr[1].ToString());
                cbSales02.Items.Add(dr[0].ToString() + "-" + dr[1].ToString());                
            }
            cbSales01.SelectedIndex = 0;
            cbSales02.SelectedIndex = 0;
            dt = null;
        }

        private void SetCustomerGrid1()
        {
            string sales01 = (cbSales01.SelectedValue.ToString().Split('-'))[0].ToString();           
            DataTable dt = CustomerBLL.getCustomerBySales(sales01);
            dgSales01.ItemsSource = dt.DefaultView;
            dt = null;
        }

        private void SetCustomerGrid2()
        {
            string sales02 = (cbSales02.SelectedValue.ToString().Split('-'))[0].ToString();
            DataTable dt = CustomerBLL.getCustomerBySales(sales02);
            dgSales02.ItemsSource = dt.DefaultView;
            dt = null;
        }

        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            if (tbSalesId.Text.Trim() != "" && tbSalesName.Text.Trim() != "")
            {
                SalesBEL salesBEL = new SalesBEL();
                salesBEL.Id = tbSalesId.Text.Trim();
                salesBEL.Name = tbSalesName.Text.Trim();
                salesBEL.IsEnable = "1";
                SalesBLL.insertSales(salesBEL);
                SetCurrentSales();
            }
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            tbSalesId.Text = "";
            tbSalesName.Text = "";
        }

        private void btnDeleteSsales_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                DataRowView rowview = dgCurrentSales.SelectedItem as DataRowView;
                string salesId = rowview.Row["Id"].ToString();
                DataTable dt = CustomerBLL.getCustomerBySales(salesId);
                if (dt.Rows.Count == 0)
                {
                    SalesBLL.updateSalesIsEnable(salesId, "0");
                    SetCurrentSales();
                }
                else
                    System.Windows.Forms.MessageBox.Show("該業務尚有客戶，請先轉移客戶至其他業務人員", "Message");
            }
            catch
            {
                System.Windows.Forms.MessageBox.Show("請先選擇要刪除的業務人員", "Message");
            }
        }

        private void cbSales01_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                SetCustomerGrid1();
            }
            catch
            {
            }
        }

        private void cbSales02_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {            
                SetCustomerGrid2();
            }
            catch
            {

            }
        }

        private void btnMoveRight_Click(object sender, RoutedEventArgs e)
        {            
            try
            {
                string salesId = (cbSales02.SelectedValue.ToString().Split('-'))[0].ToString();
                foreach (DataRowView rowView in dgSales01.SelectedItems)
                {
                    if (rowView != null)
                    {
                        DataRow row = rowView.Row;
                        string customerId = row.ItemArray[0].ToString();
                        CustomerBLL.updateResponsibilitySales(customerId, salesId);
                        
                    }
                }
                SetCustomerGrid1();
                SetCustomerGrid2();
            }
            catch
            {
                System.Windows.Forms.MessageBox.Show("請選擇要轉換的客戶", "Message");                
            }
        }

        private void btnMoveLeft_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string salesId = (cbSales01.SelectedValue.ToString().Split('-'))[0].ToString();
                foreach (DataRowView rowView in dgSales02.SelectedItems)
                {
                    if (rowView != null)
                    {
                        DataRow row = rowView.Row;
                        string customerId = row.ItemArray[0].ToString();
                        CustomerBLL.updateResponsibilitySales(customerId, salesId);

                    }
                }
                SetCustomerGrid1();
                SetCustomerGrid2();
            }
            catch
            {
                System.Windows.Forms.MessageBox.Show("請選擇要轉換的客戶", "Message");                
            }
        }
    }
}
