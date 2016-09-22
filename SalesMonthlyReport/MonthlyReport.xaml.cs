using System;
using System.Data;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Data.SqlClient;
using System.Configuration;
using SalesMonthlyReport.AppCode;
using SalesMonthlyReport.AppCode.BLL;
using SalesMonthlyReport.AppCode.BEL;


namespace SalesMonthlyReport
{
    public partial class MonthlyReport : Page
    {
        public MonthlyReport()
        {
            InitializeComponent();
            SetReportYearMonth();
        }

        private void btnImportPo_Click(object sender, RoutedEventArgs e)
        {
            string file = FunctionClass.GetExcelFile();
            if (file.Length == 0)
                System.Windows.Forms.MessageBox.Show("請選擇檔案", "Message");
            else
                ImportExcel(file, "po");            
        }

        private void btnImportFo_Click(object sender, RoutedEventArgs e)
        {
            string file = FunctionClass.GetExcelFile();
            if (file.Length == 0)
                System.Windows.Forms.MessageBox.Show("請選擇檔案", "Message");
            else
                ImportExcel(file, "");
        }

        private void btnExportReport_Click(object sender, RoutedEventArgs e)
        {
            // Create report
        }

        private void btnSave1_Click(object sender, RoutedEventArgs e)
        {
            // Update customer
            CustomerBEL objBEL = new CustomerBEL();
            foreach (DataRowView rowView in dgCustomerData1.Items)
            {
                if (rowView != null)
                {
                    DataRow row = rowView.Row;
                    objBEL.Id = row.ItemArray[0].ToString();
                    objBEL.Name = row.ItemArray[1].ToString();
                    objBEL.SalesId = (((ComboBox)row.ItemArray[2]).SelectedValue.ToString().Split('-'))[0].ToString();
                    objBEL.CountryId = "";
                    objBEL.IsEnable = "1";
                    CustomerBLL.insertCustomer(objBEL);
                }
            }
        }

        private void btnSave2_Click(object sender, RoutedEventArgs e)
        {
            // Update ForecastOrder set SalesId and CustomerId          
            foreach (DataRowView rowView in dgCustomerData1.Items)
            {
                if (rowView != null)
                {
                    DataRow row = rowView.Row;
                    string salesId = (((ComboBox)row.ItemArray[2]).SelectedValue.ToString().Split('-'))[0].ToString();
                    string customerId = row.ItemArray[0].ToString();
                    string customerName = row.ItemArray[1].ToString();       
                    ForecastOrderBLL.updateForecastOrder(salesId, customerId, customerName);
                }
            }
        }

        private void ImportExcel(string file, string type)
        {
            DataTable dt = new DataTable();
            string tableName = "";
            if (type == "po")
            {
                tableName = "tmpPurchaseOrder";
                dt = GetPurchaseOrderExcel(file);
                PurchaseOrderBLL.deleteAllTmpPurchaseOrder();
            }
            else
            {
                tableName = "tmpForecastOrder";
                dt = GetForecastOrderExcel(file);
                ForecastOrderBLL.deleteAllTmpForecastOrder();
            }

            if (dt.Rows.Count > 0)
            {
                try
                {
                    // Insert into tmp table
                    using (SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString))
                    {
                        con.Open();
                        using (SqlBulkCopy copy = new SqlBulkCopy(con))
                        {
                            int columnCount = dt.Columns.Count;
                            for (int i = 0; i < columnCount; i++)
                            {
                                copy.ColumnMappings.Add(i, i);
                            }
                            copy.DestinationTableName = tableName;
                            copy.WriteToServer(dt);
                        }
                        con.Close();
                    }
                    System.Windows.Forms.MessageBox.Show("資料匯入完畢，共匯入 " + dt.Rows.Count + " 筆資料", "Message");
                }
                catch (Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show("檔案匯入資料庫失敗", "Message");
                }
                finally
                {
                    dt = null;
                }
            }
            else
            { 
            }            
           
            if (type == "po")
            {
                // This T-SQL do :
                // 1.Copy tmp data to real table, tmpPurchaseOrder => PurchaseOrder
                // 2.Insert new customer, sales, product
                PurchaseOrderBLL.updateBasicDataByPo();
            }
            else
            {
                // This T-SQL do :
                // 1.copy tmp data to real table, tmpForecastOrder => ForecastOrder
                // 2.insert new customer, sales
                ForecastOrderBLL.updateBasicDataByFo();
            }
        }

        private DataTable GetPurchaseOrderExcel(string file)
        {
            string tableName = "[Sheet1$A6:Z]";
            string sqlCmd = @"select 銷貨單號 as Invoice, 銷貨日期 as InvoiceDate, 銷貨年度 as InvoiceYear, " +
                "客戶編號 as CustomerId, 客戶名稱 as CustomerName, 客戶訂單單號 as CustomerPurchaseOrder, 訂單日期 as OrderDate, 業務人員編號 as SalesId, 業務員 as SalesName, " +
                "產品編號 as ProductId, 客戶品號 as CustomerProductId, 產品名稱 as ProductName, 規格明細 as SpecDetail, 內箱數 as InnerBox, 外箱數 as OuterBox, " + 
                "包裝方式 as Package, 產品包裝說明 as PackageDetail, 單位 as Unit, 銷貨數量 as Quantity, 幣別 as myCurrency, 匯率 as Rate, " +
                "[單價(原)] as UnitPrice1, 單價 as UnitPrice2, [應收金額(原)] as Total1, 應收金額 as Total2, 應收稅額 as Tax from " + tableName;
            return FunctionClass.ExcelToDataTable(sqlCmd, file);
        }

        private DataTable GetForecastOrderExcel(string file)
        {
            string tableName = "[Sheet1$A6:Z]";
            string sqlCmd = @"select 預交日 as ForecastDate, 受訂日 as OrderDate, '' as SalesId, '' as CustomerId, " +            
                "客戶簡稱 as CustomerName, 客戶單號 as CustomerPurchaseOrder, 未交金額 as Total1, 訂單金額 as Total2, '" + 
                lbReportYear.Content + lbReportMonth.Content + "' as ReportYearMonth from " + tableName + " where 未交金額 <> 0";
            return FunctionClass.ExcelToDataTable(sqlCmd, file);
        }

        private void SetReportYearMonth()
        {
            int currentYear = DateTime.Now.Year;
            int currentMonth = DateTime.Now.Month;
            for (int i = 1; i <= 12; i++)
            {
                cbMonth.Items.Add(i);
            }
            lbReportYear.Content = currentYear;

            if (currentMonth >= 7 && currentMonth <= 12)
                lbReportYear.Content = currentYear + 1;
            if (currentMonth == 1)
            {
                lbReportMonth.Content = "12";
                cbMonth.SelectedIndex = 11;
            }
            else
            {
                lbReportMonth.Content = currentMonth - 1;
                cbMonth.SelectedIndex = currentMonth - 2;
            }
        }

        private void SetCustomerDg1()
        {
            // Get customers that has no sales assign to from Customer table
            DataTable dt = CustomerBLL.getCustomerWithNoSales();
            if (dt.Rows.Count != 0)
            {
                dgCustomerData1.ItemsSource = dt.DefaultView;

                // Set sales dropdown in datagrid
                dt = SalesBLL.getSalesAll();
                var Sales = new List<string>();
                Sales.Add("==請選擇業務==");
                foreach (DataRow dr in dt.Rows)
                    Sales.Add(dr[0].ToString() + "-" + dr[1].ToString());
                dgCbSales1.ItemsSource = Sales;
                btnSave1.IsEnabled = true;
            }
            else
                btnSave1.IsEnabled = false;
                       
            dt = null;
        }

        private void SetCustomerDg2()
        {
            // Get customers that has no sales assign to from ForecastOrder table                        
            DataTable dt = ForecastOrderBLL.getCustomerWithNoSalesFo();
            if (dt.Rows.Count != 0)
            {
                dgCustomerData1.ItemsSource = dt.DefaultView;

                // Set sales dropdown in datagrid
                dt = SalesBLL.getSalesAll();
                var Sales = new List<string>();
                Sales.Add("==請選擇業務==");
                foreach (DataRow dr in dt.Rows)
                    Sales.Add(dr[0].ToString() + "-" + dr[1].ToString());
                dgCbSales2.ItemsSource = Sales;
                btnSave2.IsEnabled = true;
            }
            else
                btnSave2.IsEnabled = false;

            dt = null;            
        }

        // 以下匯入初始資料用 //
        private void btnImportBasicData_Click(object sender, RoutedEventArgs e)
        {
            string file = FunctionClass.GetExcelFile();
            if (file.Length == 0)
                System.Windows.Forms.MessageBox.Show("請選擇檔案", "Message");
            else
                ImportBsicData(file);
        }

        private void ImportBsicData(string file)
        {
            string[] fileSplit = file.Split('\\');
            string ExcelTableName = "[Sheet1$A1:Z]";
            string[] sqlTableName = (fileSplit[fileSplit.Length - 1]).Split('.');
            string sqlCmd = "";
            DataTable dt = new DataTable();
            sqlCmd = "select * from " + ExcelTableName;
            dt = FunctionClass.ExcelToDataTable(sqlCmd, file);

            if (dt.Rows.Count > 0)
            {
                try
                {
                    using (SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString))
                    {
                        con.Open();
                        using (SqlBulkCopy copy = new SqlBulkCopy(con))
                        {
                            int columnCount = dt.Columns.Count;
                            for (int i = 0; i < columnCount; i++)
                            {
                                copy.ColumnMappings.Add(i, i);
                            }
                            copy.DestinationTableName = sqlTableName[0].ToString();
                            copy.WriteToServer(dt);                            
                        }
                        con.Close();
                    }
                    System.Windows.Forms.MessageBox.Show("Done", "Message");
                }
                catch (Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.ToString(), "Message");
                }
                finally
                {
                    dt = null;
                }
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("No data row found in excel file.", "Message");
            }
        }
    }
}
