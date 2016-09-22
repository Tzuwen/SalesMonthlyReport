using System;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using SalesMonthlyReport.AppCode.BEL;

namespace SalesMonthlyReport.AppCode.DAL
{
    public class SalesDAL
    {
        SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["DefaultConnection"].ToString());
        
        public Int32 insertSales(SalesBEL objBEL)
        {
            int result;
            try
            {
                SqlCommand cmd = new SqlCommand("dbo.sp_insertSales", con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Id", objBEL.Id);
                cmd.Parameters.AddWithValue("@Name", objBEL.Name);
                cmd.Parameters.AddWithValue("@IsEnable", objBEL.IsEnable);

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

        public Int32 updateSalesIsEnable(string id, string isEnable)
        {
            int result;
            try
            {
                SqlCommand cmd = new SqlCommand("dbo.sp_updateSalesIsEnable", con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Id", id);               
                cmd.Parameters.AddWithValue("@IsEnable", isEnable);

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

        public DataTable getSalesAll()
        {
            DataSet ds = new DataSet();
            try
            {
                SqlCommand cmd = new SqlCommand("dbo.sp_selectSalesAll", con);
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
