using System;
using System.Data;
using SalesMonthlyReport.AppCode.DAL;
using SalesMonthlyReport.AppCode.BEL;

namespace SalesMonthlyReport.AppCode.BLL
{
    public class SalesBLL
    {
        public static Int32 insertSales(SalesBEL objBel)
        {
            SalesDAL objDal = new SalesDAL();
            try
            {
                return objDal.insertSales(objBel);
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

        public static Int32 updateSalesIsEnable(string id, string isEnable)
        {
            SalesDAL objDal = new SalesDAL();
            try
            {
                return objDal.updateSalesIsEnable(id, isEnable);
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

        public static DataTable getSalesAll()
        {
            SalesDAL objDal = new SalesDAL();
            try
            {
                return objDal.getSalesAll();
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
