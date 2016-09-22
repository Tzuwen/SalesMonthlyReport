using System;
using System.Data;
using System.Data.OleDb;

namespace SalesMonthlyReport.AppCode
{
    class FunctionClass
    {
        public static string GetExcelFile()
        {
            string fileName = "";
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.DefaultExt = ".xls";
            dlg.Filter = "Excel Files (*.xls, *.xlsx)|*.xls;*.xlsx";
            bool? result = dlg.ShowDialog();
            if (result == true)
                fileName = dlg.FileName;
            return fileName;
        }

        public static DataTable ExcelToDataTable(string sql, string file)
        {
            DataTable dt = new DataTable();
            try
            {
                OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + file + ";Extended Properties='Excel 12.0 Xml;HDR=YES'");
                OleDbDataAdapter da = new OleDbDataAdapter(sql, conn);
                da.Fill(dt);
                dt.TableName = "tmp";
                conn.Close();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Excel讀取失敗，請檢查Excel格式", "Message");
            }
            finally
            {
            }
            return dt;
        }
    }
}
