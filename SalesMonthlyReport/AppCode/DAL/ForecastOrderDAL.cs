using System;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;

namespace SalesMonthlyReport.AppCode.DAL
{
    public class ForecastOrderDAL
    {
        SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString);

        public Int32 updateForecastOrder(string salesId, string customerId, string customerName)
        {
            int result;
            try
            {
                SqlCommand cmd = new SqlCommand("dbo.sp_updateForecastOrder", con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@SalesId", salesId);
                cmd.Parameters.AddWithValue("@CustomerId", customerId);
                cmd.Parameters.AddWithValue("@CustomerName", customerName);
                if (con.State == ConnectionState.Closed)
                    con.Open();

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
                if (con.State != ConnectionState.Closed)
                    con.Close();
            }
        }

        public Int32 updateBasicDataByFo()
        {
            int result;
            try
            {
                SqlCommand cmd = new SqlCommand("dbo.sp_updateBasicDataByFo", con);
                cmd.CommandType = CommandType.StoredProcedure;

                if (con.State == ConnectionState.Closed)
                    con.Open();

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
                if (con.State != ConnectionState.Closed)
                    con.Close();
            }
        }

        public Int32 deleteAllTmpForecastOrder()
        {
            int result;
            try
            {
                SqlCommand cmd = new SqlCommand("dbo.sp_deleteAllTmpForecastOrder", con);
                cmd.CommandType = CommandType.StoredProcedure;

                if (con.State == ConnectionState.Closed)
                    con.Open();

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
                if (con.State != ConnectionState.Closed)
                    con.Close();
            }
        }

        public DataTable getCustomerWithNoSalesFo()
        {
            DataSet ds = new DataSet();
            try
            {
                SqlCommand cmd = new SqlCommand("dbo.sp_selectCustomerWithNoSalesFo", con);
                cmd.CommandType = CommandType.StoredProcedure;
                SqlDataAdapter adp = new SqlDataAdapter(cmd);
                adp.Fill(ds);
                cmd.Dispose();
            }
            catch (Exception ex)
            {
                throw;
            }
            finally
            {
                ds.Dispose();
            }
            return ds.Tables[0];
        }
    }
}
