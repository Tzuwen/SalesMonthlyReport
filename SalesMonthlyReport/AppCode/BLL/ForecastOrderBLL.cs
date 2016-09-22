using System;
using System.Data;
using SalesMonthlyReport.AppCode.DAL;

namespace SalesMonthlyReport.AppCode.BLL
{
    public class ForecastOrderBLL
    {
        public static Int32 updateForecastOrder(string salesId, string customerId, string customerName)
        {
            ForecastOrderDAL objDal = new ForecastOrderDAL();
            try
            {
                return objDal.updateForecastOrder(salesId, customerId, customerName);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                objDal = null;
            }
        }

        public static Int32 updateBasicDataByFo()
        {
            ForecastOrderDAL objDal = new ForecastOrderDAL();
            try
            {
                return objDal.updateBasicDataByFo();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                objDal = null;
            }
        }

        public static Int32 deleteAllTmpForecastOrder()
        {
            ForecastOrderDAL objDal = new ForecastOrderDAL();
            try
            {
                return objDal.deleteAllTmpForecastOrder();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                objDal = null;
            }
        }

        public static DataTable getCustomerWithNoSalesFo()
        {
            ForecastOrderDAL objDal = new ForecastOrderDAL();
            try
            {
                return objDal.getCustomerWithNoSalesFo();
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
