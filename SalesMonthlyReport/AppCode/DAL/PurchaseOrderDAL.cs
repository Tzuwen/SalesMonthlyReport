using System;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;

namespace SalesMonthlyReport.AppCode.DAL
{
    public class PurchaseOrderDAL
    {
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString);

        public Int32 updateBasicDataByPo()
        {
            int result;
            try
            {
                SqlCommand cmd = new SqlCommand("dbo.sp_updateBasicDataByPo", conn);
                cmd.CommandType = CommandType.StoredProcedure;

                if (conn.State == ConnectionState.Closed)
                    conn.Open();

                result = cmd.ExecuteNonQuery();
                cmd.Dispose();

                if (result > 0)
                    return result;
                else
                    return 0;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (conn.State != ConnectionState.Closed)
                    conn.Close();
            }
        }

        public Int32 deleteAllTmpPurchaseOrder()
        {
            int result;
            try
            {
                SqlCommand cmd = new SqlCommand("dbo.sp_deleteAllTmpPurchaseOrder", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                             
                if (conn.State == ConnectionState.Closed)
                    conn.Open();

                result = cmd.ExecuteNonQuery();
                cmd.Dispose();

                if (result > 0)
                    return result;
                else
                    return 0;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (conn.State != ConnectionState.Closed)
                    conn.Close();
            }
        }
    }
}
