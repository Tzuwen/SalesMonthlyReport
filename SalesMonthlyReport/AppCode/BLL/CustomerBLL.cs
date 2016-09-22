using System;
using System.Data;
using SalesMonthlyReport.AppCode.DAL;
using SalesMonthlyReport.AppCode.BEL;

namespace SalesMonthlyReport.AppCode.BLL
{
    public class CustomerBLL
    {
        public static Int32 insertCustomer(CustomerBEL objBel)
        {
            CustomerDAL objDal = new CustomerDAL();
            try
            {
                return objDal.insertCustomer(objBel);
            }
            catch (Exception ex)
            {
                throw;
            }
            finally
            {
                objDal = null;
            }
        }

        public static Int32 updateResponsibilitySales(string id, string salesId)
        {
            CustomerDAL objDal = new CustomerDAL();
            try
            {
                return objDal.updateResponsibilitySales(id, salesId);
            }
            catch (Exception ex)
            {
                throw;
            }
            finally
            {
                objDal = null;
            }
        }

        public static DataTable getCustomerBySales(string salesId)
        {
            CustomerDAL objDal = new CustomerDAL();
            try
            {
                return objDal.getCustomerBySales(salesId);
            }
            catch (Exception ex)
            {
                throw;
            }
            finally
            {
                objDal = null;
            }
        }

        public static DataTable getCustomerWithNoSales()
        {
            CustomerDAL objDal = new CustomerDAL();
            try
            {
                return objDal.getCustomerWithNoSales();
            }
            catch (Exception ex)
            {
                throw;
            }
            finally
            {
                objDal = null;
            }
        }
    }
}
