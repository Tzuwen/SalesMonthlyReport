using System;
using System.Data;
using SalesMonthlyReport.AppCode.DAL;
using SalesMonthlyReport.AppCode.BEL;


namespace SalesMonthlyReport.AppCode.BLL
{
    public class YearForecastBLL
    {
        public static Int32 insertCustomer(YearForecastBEL objBel)
        {
            YearForecastDAL objDal = new YearForecastDAL();
            try
            {
                return objDal.insertYearForecast(objBel);
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
       
        public static DataTable getYearForecastByYearSales(string year, string salesId)
        {
            YearForecastDAL objDal = new YearForecastDAL();
            try
            {
                return objDal.getYearForecastByYearSales(year, salesId);
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
