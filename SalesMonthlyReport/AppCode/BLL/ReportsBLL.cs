using System;
using System.Data;
using SalesMonthlyReport.AppCode.DAL;

namespace SalesMonthlyReport.AppCode.BLL
{
    public class ReportsBLL
    {
        public static Int32 calculateReport(int year)
        {
            ReportsDAL objDal = new ReportsDAL();
            try
            {
                return objDal.calculateReport(year);
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

        public static DataTable getSalesReportS1T1()
        {
            ReportsDAL objDal = new ReportsDAL();
            try
            {
                return objDal.getSalesReportS1T1();
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

        public static DataTable getSalesReportS1T2()
        {
            ReportsDAL objDal = new ReportsDAL();
            try
            {
                return objDal.getSalesReportS1T2();
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

        public static DataTable getSalesReportS1T3()
        {
            ReportsDAL objDal = new ReportsDAL();
            try
            {
                return objDal.getSalesReportS1T3();
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

        public static DataTable getSalesReportS2()
        {
            ReportsDAL objDal = new ReportsDAL();
            try
            {
                return objDal.getSalesReportS2();
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
        
        public static DataTable getSalesPersonalReport(string salesId)
        {
            ReportsDAL objDal = new ReportsDAL();
            try
            {
                return objDal.getSalesPersonalReport(salesId);
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

        public static DataTable getSalesPersonalReportTotal()
        {
            ReportsDAL objDal = new ReportsDAL();
            try
            {
                return objDal.getSalesPersonalReportTotal();
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

        public static DataTable getSalesPersonalForecast(string salesId)
        {
            ReportsDAL objDal = new ReportsDAL();
            try
            {
                return objDal.getSalesPersonalForecast(salesId);
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

        public static DataTable getSalesPersonalForecastTotal()
        {
            ReportsDAL objDal = new ReportsDAL();
            try
            {
                return objDal.getSalesPersonalForecastTotal();
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
    }
}
