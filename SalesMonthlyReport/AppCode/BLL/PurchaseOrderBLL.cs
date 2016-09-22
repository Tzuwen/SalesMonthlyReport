using System;
using SalesMonthlyReport.AppCode.DAL;

namespace SalesMonthlyReport.AppCode.BLL
{
    public class PurchaseOrderBLL
    {
        public static Int32 updateBasicDataByPo()
        {
            PurchaseOrderDAL objDal = new PurchaseOrderDAL();
            try
            {
                return objDal.updateBasicDataByPo();
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

        public static Int32 deleteAllTmpPurchaseOrder()
        {
            PurchaseOrderDAL objDal = new PurchaseOrderDAL();
            try
            {
                return objDal.deleteAllTmpPurchaseOrder();
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
