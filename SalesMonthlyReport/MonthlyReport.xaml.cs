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
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.IO;
using System.Collections;

namespace SalesMonthlyReport
{
    public partial class MonthlyReport : Page
    {
        public MonthlyReport()
        {
            InitializeComponent();
            SetReportYearMonth();
            SetCustomerDg1();
            SetCustomerDg2();
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
            ReportsBLL.calculateReport(Convert.ToInt32(this.lbReportYear.Content));

            // Insert data into excel sheet
            using (ExcelPackage p = new ExcelPackage())
            {
                // Here setting some document properties
                p.Workbook.Properties.Author = "YM";
                p.Workbook.Properties.Title = "Sales Monthly Report";

                // Create a sheet
                p.Workbook.Worksheets.Add("WorkSheet");
                ExcelWorksheet ws = p.Workbook.Worksheets[1];


                // Write to excel sheet 1
                // sheet 1 table 1 = s1t1, sheet 1 table 2 = s1t2, sheet 1 table 3 = s1t3
                int currentRow = 1;
                currentRow = writeToExcel("s1t1", currentRow, ws, "");
                currentRow = writeToExcel("s1t2", currentRow, ws, "");
                writeToExcel("s1t3", currentRow, ws, "");
                // Write to excel sheet 2 
                currentRow = 1;
                p.Workbook.Worksheets.Add("WorkSheet2");
                ExcelWorksheet ws2 = p.Workbook.Worksheets[2];
                p.Workbook.Worksheets.Add("WorkSheet3");
                ExcelWorksheet ws3 = p.Workbook.Worksheets[3];
                DataTable dt = SalesBLL.getSalesAll();
                if (dt.Rows.Count != 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        currentRow = writeToExcel("s2", currentRow, ws2, dr["Id"].ToString());                       
                    }
                    // Write s2 Total
                    writeToExcel("s2", currentRow + 2, ws2, "");

                    currentRow = 1;
                    foreach (DataRow dr in dt.Rows)
                    {
                        currentRow = writeToExcel("s3", currentRow, ws3, dr["Id"].ToString());
                    }
                    // Write s3 Total
                    writeToExcel("s3", currentRow + 2, ws3, "");
                }

                // Generate A File with Random name
                Byte[] bin = p.GetAsByteArray();
                //string file = "d:\\" + Guid.NewGuid().ToString() + ".xlsx";
                string savePath = saveFile();
                string file = savePath;
                File.WriteAllBytes(file, bin);
                System.Windows.Forms.MessageBox.Show("檔案儲存完畢", "Message");

            }
        }

        private void btnSave1_Click(object sender, RoutedEventArgs e)
        {
            // Update customer
            string[] sales = null;
            string salesId = "";
            CustomerBEL objBEL = new CustomerBEL();
            for (int i = 0; i < dgCustomerData1.Items.Count; i++)
            {
                DataGridCell cell0 = FunctionClass.GetCell(dgCustomerData1, i, 0);
                DataGridCell cell1 = FunctionClass.GetCell(dgCustomerData1, i, 1);
                DataGridCell cell2 = FunctionClass.GetCell(dgCustomerData1, i, 2);
                TextBlock tb0 = cell0.Content as TextBlock;
                TextBlock tb1 = cell1.Content as TextBlock;
                ComboBox cb = cell2.Content as ComboBox;
                objBEL.Id = tb0.Text.Trim();
                objBEL.Name = tb1.Text.Trim();
                objBEL.CountryId = "";
                objBEL.IsEnable = "1";
                if (cb.SelectedValue != null)
                {
                    sales = cb.SelectedValue.ToString().Split('-');
                    salesId = sales[0].ToString().Trim();
                    CustomerBLL.insertCustomer(objBEL);
                }
            }
        }

        private void btnSave2_Click(object sender, RoutedEventArgs e)
        {
            // Update ForecastOrder set SalesId and CustomerId
            string customerId = "";
            string customerName = "";           
            string[] sales = null;
            string salesId = "";
            CustomerBEL objBEL = new CustomerBEL();
            for (int i = 0; i< dgCustomerData2.Items.Count; i++)
            {
                DataGridCell cell0 = FunctionClass.GetCell(dgCustomerData2, i, 0);
                DataGridCell cell1 = FunctionClass.GetCell(dgCustomerData2, i, 1);
                DataGridCell cell2 = FunctionClass.GetCell(dgCustomerData2, i, 2);
                TextBlock tb0 = cell0.Content as TextBlock;
                TextBlock tb1 = cell1.Content as TextBlock;
                ComboBox cb = cell2.Content as ComboBox;
                customerId = tb0.Text.Trim();
                customerName = tb1.Text.Trim();
                if (cb.SelectedValue != null)
                {
                    sales = cb.SelectedValue.ToString().Split('-');
                    salesId = sales[0].ToString().Trim();
                    ForecastOrderBLL.updateForecastOrder(salesId, customerId, customerName);
                    objBEL.Id = customerId;
                    objBEL.Name = customerName;
                    objBEL.CountryId = "";
                    objBEL.IsEnable = "1";
                    objBEL.SalesId = salesId;
                    CustomerBLL.insertCustomer(objBEL);
                }
            }
        }
      
        private IEnumerable<DataGridRow> getDataGridRows(DataGrid grid)
        {
            var itemsSource = grid.ItemsSource as IEnumerable;
            if (null == itemsSource) yield return null;
            foreach (var item in itemsSource)
            {
                var row = grid.ItemContainerGenerator.ContainerFromItem(item) as DataGridRow;                
                if (null != row) yield return row;
            }
        }

        // insert using SqlBulkCopy
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
                SetCustomerDg1();
            }
            else
            {
                // This T-SQL do :
                // 1.copy tmp data to real table, tmpForecastOrder => ForecastOrder
                // 2.insert new customer, sales
                ForecastOrderBLL.updateBasicDataByFo();
                SetCustomerDg2();
            }
        }

        private int writeToExcel(string type, int startRow, ExcelWorksheet ws, string salesId)
        {            
            DataTable dt = new DataTable();
            int result = startRow;
            List<string> titles = new List<string>();

            ws.Cells.Style.Font.Size = 11; // Default font size for whole sheet
            ws.Cells.Style.Font.Name = "Calibri"; // Default Font name for whole sheet            

            var reportYear = Convert.ToInt16(lbReportYear.Content);
            var reportMonth = Convert.ToInt16(lbReportMonth.Content);
            var startYearMonth1 = (reportYear - 1).ToString() + "/07";
            var currentYearMonth1 = reportMonth >= 7 ? (reportYear - 1).ToString() + "/" + reportMonth.ToString() : (reportYear).ToString() + "/" + reportMonth.ToString();

            var startYearMonth2 = (reportYear - 2).ToString() + "/07";
            var currentYearMonth2 = reportMonth >= 7 ? (reportYear - 2).ToString() + "/" + reportMonth.ToString() : (reportYear-1).ToString() + "/" + reportMonth.ToString();

            var startYearMonth3 = (reportYear - 3).ToString() + "/07";
            var currentYearMonth3 = reportMonth >= 7 ? (reportYear - 3).ToString() + "/" + reportMonth.ToString() : (reportYear-2).ToString() + "/" + reportMonth.ToString();

            // Set Title for diffrent table
            switch (type)
            {
                case "s1t1":
                case "s1t3":
                    ws.Name = "年度比較"; // Setting Sheet's name
                    titles.Add("業務員");
                    titles.Add(reportYear.ToString() + "年度" + "(" + startYearMonth1 + "~" + currentYearMonth1 + ")");
                    titles.Add((reportYear - 1).ToString() + "年度" + "(" + startYearMonth2 + "~" + currentYearMonth2 + ")");
                    titles.Add((reportYear - 2).ToString() + "年度" + "(" + startYearMonth2 + "~" + currentYearMonth3 + ")");
                    titles.Add((reportYear - 1).ToString() + "v.s" + reportYear.ToString() + "成長率");
                    titles.Add((reportYear - 2).ToString() + "v.s" + reportYear.ToString() + "成長率");
                    if (type == "s1t1")
                    {
                        ws.Cells[startRow, 1].Value = "同期間比較(已銷+未出)";
                        dt = ReportsBLL.getSalesReportS1T1();
                    }
                    else
                    {
                        ws.Cells[startRow, 1].Value = "同期間比較(已銷不含未出)";
                        dt = ReportsBLL.getSalesReportS1T3();
                    }                    
                    break;
                case "s1t2":
                    titles.Add("業務員");
                    titles.Add(reportYear.ToString() + "年度已銷+未出(累計至" + currentYearMonth1 + ")");
                    titles.Add((reportYear - 1).ToString() + "銷貨金額");
                    titles.Add((reportYear - 2).ToString() + "銷貨金額");
                    titles.Add(reportYear.ToString() + "達成率");
                    titles.Add(reportYear.ToString() + "預估");
                    titles.Add("預估與實際達成率");
                    ws.Cells[startRow, 1].Value = "累計 vs 銷售";
                    dt = ReportsBLL.getSalesReportS1T2();
                    break;
                case "s2":
                    ws.Name = "當月、年度比較"; // Setting Sheet's name
                    titles.Add("業務員");
                    titles.Add("客戶編號");
                    titles.Add("客戶名稱");
                    titles.Add(reportYear.ToString() + "年度已銷(" + startYearMonth1 + "~" + currentYearMonth1  + ")");
                    titles.Add((reportYear-1).ToString() + "年度已銷(" + startYearMonth2 + "~" + currentYearMonth2 + ")");
                    titles.Add((reportYear-2).ToString() + "年度已銷(" + startYearMonth3 + "~" + currentYearMonth3 + ")");
                    titles.Add((reportYear - 1).ToString() + "v.s" + reportYear.ToString() + "同期成長率((D/E)-1)");
                    titles.Add(reportYear.ToString() + "年度已接未出貨(" + reportYear.ToString() + "/06/30)");
                    titles.Add(reportYear.ToString() + "年度已銷+未出(" + startYearMonth1 + "~" + currentYearMonth1 + ")(D+H)");
                    titles.Add((reportYear-1).ToString() + "年度已銷+未出(" + startYearMonth2 + "~" + currentYearMonth2 + ")");
                    titles.Add((reportYear-2).ToString() + "年度已銷+未出(" + startYearMonth3 + "~" + currentYearMonth3 + ")");
                    titles.Add((reportYear - 1).ToString() + "v.s" + reportYear.ToString() + "成長率((I/J)-1)");
                    titles.Add((reportYear - 2).ToString() + "v.s" + reportYear.ToString() + "成長率((I/K)-1)");
                    titles.Add((reportYear - 1).ToString() + "總銷貨額");
                    titles.Add((reportYear - 2).ToString() + "總銷貨額");                   
                    titles.Add(reportYear.ToString() + "年度達成率(I/N)");
                    titles.Add(reportYear.ToString() + "預估");
                    titles.Add("預估與實際達成率(I/Q)");
                    titles.Add(reportYear.ToString() + "/" + reportMonth.ToString() + "下單金額");
                    if (salesId != "")
                    {
                        dt = ReportsBLL.getSalesPersonalReport(salesId);
                    }
                    else
                    {
                        dt = ReportsBLL.getSalesPersonalReportTotal();
                    }
                    break;
                case "s3":
                    ws.Name = "每月預估出貨金額"; // Setting Sheet's name
                    titles.Add("業務員");
                    titles.Add("客戶編號");
                    titles.Add("客戶名稱");
                    titles.Add((reportYear - 1).ToString()+".07");
                    titles.Add((reportYear - 1).ToString() + ".08");
                    titles.Add((reportYear - 1).ToString() + ".09");
                    titles.Add((reportYear - 1).ToString() + ".10");
                    titles.Add((reportYear - 1).ToString() + ".11");
                    titles.Add((reportYear - 1).ToString() + ".12");
                    titles.Add((reportYear).ToString() + ".01");
                    titles.Add((reportYear).ToString() + ".02");
                    titles.Add((reportYear).ToString() + ".03");
                    titles.Add((reportYear).ToString() + ".04");
                    titles.Add((reportYear).ToString() + ".05");
                    titles.Add((reportYear).ToString() + ".06");
                    titles.Add("Total");
                    if (salesId != "")
                    {
                        dt = ReportsBLL.getSalesPersonalForecast(salesId);
                    }
                    else
                    {
                        dt = ReportsBLL.getSalesPersonalForecastTotal();
                    }
                    break;

            }

            if (dt.Rows.Count > 1)
            {
                if (type != "s2" && type != "s3")
                {
                    // Merging cells and create a center heading for out table            
                    ws.Cells[startRow, 1, startRow, dt.Columns.Count - 1].Merge = true;
                    ws.Cells[startRow, 1, startRow, dt.Columns.Count - 1].Style.Font.Bold = true;
                    //ws.Cells[startRow, 1, startRow, dt.Columns.Count-1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Column(1).Width = 18;
                    ws.Column(2).Width = 18;
                    ws.Column(3).Width = 18;
                    ws.Column(4).Width = 18;
                    ws.Column(5).Width = 12;
                    ws.Column(6).Width = 18;
                    ws.Column(7).Width = 12;
                }
                else if (type == "s2")
                {
                    ws.Column(1).Width = 12;
                    ws.Column(2).Width = 12;
                    ws.Column(3).Width = 12;
                    ws.Column(4).Width = 18;
                    ws.Column(5).Width = 18;
                    ws.Column(6).Width = 18;
                    ws.Column(7).Width = 12;
                    ws.Column(8).Width = 18;
                    ws.Column(9).Width = 18;
                    ws.Column(10).Width = 18;
                    ws.Column(11).Width = 18;
                    ws.Column(12).Width = 12;
                    ws.Column(13).Width = 12;
                    ws.Column(14).Width = 18;
                    ws.Column(15).Width = 18;
                    ws.Column(16).Width = 12;
                    ws.Column(17).Width = 18;
                    ws.Column(18).Width = 12;
                    ws.Column(19).Width = 18;
                }
                else
                {
                    ws.Column(1).Width = 12;
                    ws.Column(2).Width = 12;
                    ws.Column(3).Width = 12;
                    ws.Column(4).Width = 14;
                    ws.Column(5).Width = 14;
                    ws.Column(6).Width = 14;
                    ws.Column(7).Width = 14;
                    ws.Column(8).Width = 14;
                    ws.Column(9).Width = 14;
                    ws.Column(10).Width = 14;
                    ws.Column(11).Width = 14;
                    ws.Column(12).Width = 14;
                    ws.Column(13).Width = 14;
                    ws.Column(14).Width = 14;
                    ws.Column(15).Width = 14;
                    ws.Column(16).Width = 14;
                    ws.Column(17).Width = 14;
                }

                int colIndex = 1;
                int rowIndex = startRow + 1;

                // Create table header
                if (type == "s2" && startRow != 1 || type == "s3" && startRow != 1)
                { }
                else
                {
                    foreach (var title in titles)
                    {
                        var cell = ws.Cells[rowIndex, colIndex];
                        cell.Style.WrapText = true;
                        // Setting Top/left,right/bottom borders.
                        var border = cell.Style.Border;
                        border.Bottom.Style =
                            border.Top.Style =
                            border.Left.Style =
                            border.Right.Style = ExcelBorderStyle.Thin;

                        // Setting Value in cell
                        cell.Value = title;
                        colIndex++;
                    }
                }

                // Insert table content
                //foreach (DataRow dr in dt.Rows) // Adding Data into rows
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    rowIndex++;
                    for (int j = 1; j < dt.Columns.Count; j++)
                    {
                        var cell = ws.Cells[rowIndex, j];                     

                        switch (type)
                        {
                            case "s1t1":
                            case "s1t3":
                                cell.Value = dt.Rows[i][j];
                                if (j >= 2 && j <= 4)
                                {
                                    cell.Style.Numberformat.Format = "$#,##0";
                                }
                                else if (j >= 5 && j <= 6)
                                {
                                    cell.Style.Numberformat.Format = "0.00%";
                                }
                                break;
                            case "s1t2":
                                cell.Value = dt.Rows[i][j];
                                if (j >= 2 && j <= 4 || j == 6)
                                {
                                    cell.Style.Numberformat.Format = "$#,##0";
                                }
                                else if (j == 5 || j == 7)
                                {
                                    cell.Style.Numberformat.Format = "0.00%";
                                }
                                break;
                            case "s2":
                                if (i != 0 && j == 1)
                                {

                                }
                                else
                                {
                                    cell.Value = dt.Rows[i][j];
                                }
                                if (j >= 4 && j <= 6 || j >= 8 && j <= 11 || j >= 14 && j <= 15 || j == 17 || j == 19)
                                {
                                    cell.Style.Numberformat.Format = "$#,##0";
                                }
                                else if (j == 7 || j >= 12 && j <= 13 || j == 16 || j == 18)
                                {
                                    cell.Style.Numberformat.Format = "0.00%";
                                }
                                break;
                            case "s3":
                                if (i != 0 && j == 1)
                                {

                                }
                                else
                                {
                                    cell.Value = dt.Rows[i][j];
                                }
                                if (j >= 4)
                                {
                                    cell.Style.Numberformat.Format = "$#,##0";
                                }
                                break;
                        }

                        // Setting borders of cell
                        var border = cell.Style.Border;
                        border.Left.Style = border.Right.Style = ExcelBorderStyle.Thin;
                        colIndex++;
                    }
                }

                if (type != "s2" && type != "s3")
                {
                    result += dt.Rows.Count + 5;
                }
                else
                {
                    result += dt.Rows.Count;
                }
            }

            dt.Clear();
            return result;
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
                dgCustomerData2.ItemsSource = dt.DefaultView;

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

        private string saveFile()
        {
            string saveFile = "";
            // Create OpenFileDialog 
            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
            // Set filter for file extension and default file extension 
            dlg.DefaultExt = ".xlsx";
            dlg.Filter = "XLSX Files (*.xlsx)|*.xlsx";
            Nullable<bool> result = dlg.ShowDialog();
            if (result == true)
            {
                saveFile = dlg.FileName;
            }
            return saveFile;
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
