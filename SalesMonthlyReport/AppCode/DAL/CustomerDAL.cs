using System;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using SalesMonthlyReport.AppCode.BEL;

namespace SalesMonthlyReport.AppCode.DAL
{
    class CustomerDAL
    {
        SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["DefaultConnection"].ToString());

        public Int32 insertCustomer(CustomerBEL objBEL)
        {
            int result;
            try
            {
                SqlCommand cmd = new SqlCommand("dbo.sp_insertCustomer", con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Id", objBEL.Id);
                cmd.Parameters.AddWithValue("@Name", objBEL.Name);
                cmd.Parameters.AddWithValue("@SalesId", objBEL.SalesId);
                cmd.Parameters.AddWithValue("@CountryId", objBEL.CountryId);
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

        public Int32 updateResponsibilitySales(string id, string salesId)
        {
            int result;
            try
            {
                SqlCommand cmd = new SqlCommand("dbo.sp_updateResponsibilitySales", con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Id", id);
                cmd.Parameters.AddWithValue("@SalesId", salesId);

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

        public DataTable getCustomerBySales(string salesId)
        {
            DataSet ds = new DataSet();
            try
            {
                SqlCommand cmd = new SqlCommand("dbo.sp_selectCustomerBySales", con);
                cmd.CommandType = CommandType.StoredProcedure;
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

        public DataTable getCustomerWithNoSales()
        {
            DataSet ds = new DataSet();
            try
            {
                SqlCommand cmd = new SqlCommand("dbo.sp_selectCustomerWithNoSales", con);
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
