using System;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using SalesMonthlyReport.AppCode.BEL;

namespace SalesMonthlyReport.AppCode.DAL
{
    class YearForecastDAL
    {
        SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["DefaultConnection"].ToString());

        public Int32 insertYearForecast(YearForecastBEL objBEL)
        {
            int result;
            try
            {
                SqlCommand cmd = new SqlCommand("dbo.sp_insertYearForecast", con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@CustomerId", objBEL.CustomerId);
                cmd.Parameters.AddWithValue("@Year", objBEL.Year);
                cmd.Parameters.AddWithValue("@Forecast", objBEL.Forecast);

                if (con.State == ConnectionState.Closed)
                {
                    con.Open();
                }
                result = cmd.ExecuteNonQuery();
                cmd.Dispose();
                if (result > 0)
                {
                    return result;
                }
                else
                {
                    return 0;
                }
            }
            catch (Exception ex)
            {
                throw;
            }
            finally
            {
                if (con.State != ConnectionState.Closed)
                {
                    con.Close();
                }
            }
        }

        public DataTable getYearForecastByYearSales(string year, string salesId)
        {
            DataSet ds = new DataSet();
            try
            {
                SqlCommand cmd = new SqlCommand("dbo.sp_selectYearForecastByYearSales", con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Year", year);
                cmd.Parameters.AddWithValue("@SalesId", salesId);
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
