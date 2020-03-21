using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PRO_Metrics_NET.Models;
using STXtoSQL.Log;
using PRO_Metrics_NET.DataAccess;

namespace PRO_Metrics_NET
{
    class Program
    {
        static void Main(string[] args)
        {
            // Args will change based on STXtoSQL program goal
            string date1 = "";
            string date2 = "";
            string brh = "";
            int odbcCnt = 0;
            int insertCnt = 0;
            int importCnt = 0;

            #region Args
            try
            {
                if (args.Length == 0)
                {                   
                    Logger.LogWrite("MSG", "No args present");
                    //return;
                    Console.WriteLine("No args present");
                }
                else if (args.Length == 1)
                {
                    // 1st arg should be the Brh
                    brh = args[0].ToString();

                    /*
                     * Next two args are date range, but not
                     * provided so use range for current
                     * month to yesterday
                     */
                    DateTime dtToday = DateTime.Today;

                    DateTime dtFirst = new DateTime(dtToday.Year, dtToday.Month, 1);

                    /*
                     * Need one date part of datetime.
                     * Time and date are separated by a space, so split the string
                     * and only use the 1st element.
                     */
                    string[] date1Split = dtFirst.ToString().Split(' ');
                    string[] date2Split = dtToday.AddDays(-1).ToString().Split(' ');

                    /*
                    * Must be in format mm/dd/yyyy.  No time part
                    */
                    date1 = date1Split[0];
                    date2 = date2Split[0];                  
                }
                else
                {
                    /*
                     * Date range provided
                     * Must be in format mm/dd/yyyy.  No time part
                     */
                    date1 = args[0].ToString();
                    date2 = args[1].ToString();
                }
            }
            catch (Exception ex)
            {
                Logger.LogWrite("EXC", ex);
                Logger.LogWrite("MSG", "Return");
                return;
                //Console.WriteLine(ex.Message.ToString());
            }
            #endregion

            #region PWC
            /*
             * Get all PWC for brh
             * PWC are created in a StratixData table.  The values
             * equal the Tabs in the Excel file, NOT the PWC in 
             * the Stratix tables
             */
            SQLData objSQL = new SQLData();

            SQLData objSQL_PWC = new SQLData();

            List<string> lstPWC = new List<string>();

            try
            {
                lstPWC = objSQL_PWC.Get_PWC_ByBrh(brh);              
            }
            catch(Exception ex)
            {
                Logger.LogWrite("EXC", ex);
                Logger.LogWrite("MSG", "Return");
                return;
            }
            #endregion

            #region Bookings
            SQLData objSQL_Bookings = new SQLData();

            List<BookingsModel> lstBookings = new List<BookingsModel>();

            try
            {
                lstBookings = objSQL_Bookings.Get_Bookings_ByBrh(brh, date1, date2);
                // testing
                Flat.WriteBookingsFlatFile(lstBookings);
            }
            catch (Exception ex)
            {
                Logger.LogWrite("EXC", ex);
                Logger.LogWrite("MSG", "Return");
                return;
            }
            #endregion

            #region Sales
            List<SalesModel> lstSales = new List<SalesModel>();

            try
            {
                lstSales = objSQL.Get_Sales_ByBrh(brh, date1, date2);
                // testing
                Flat.WriteSalesFlatFile(lstSales);
            }
            catch (Exception ex)
            {
                Logger.LogWrite("EXC", ex);
                Logger.LogWrite("MSG", "Return");
                return;
            }
            #endregion

            #region Prod Sum
            List<ProdSumModel> lstProdSum = new List<ProdSumModel>();

            try
            {
                lstProdSum = objSQL.Get_Prod_Sum(date1, date2);
                // testing
                Flat.WriteProdSumFlatFile(lstProdSum);
            }
            catch (Exception ex)
            {
                Logger.LogWrite("EXC", ex);
                Logger.LogWrite("MSG", "Return");
                return;
            }
            #endregion

            #region Prod by PWC
            /*
             * Get Production for each PWC.  Add all to one List<ProdPWCModel>.
             * Filter List with LINQ to insert into specific XLS tabs
             */
            List<ProdPWCModel> lstProdPWCAll = new List<ProdPWCModel>();

            foreach (string pwc in lstPWC)
            {
                List<ProdPWCModel> lstProdPWC = new List<ProdPWCModel>();

                try
                {
                    lstProdPWC = objSQL.Get_Prod_PWC_ByPWC(pwc, date1, date2);
                    // testing
                    Flat.WriteProdPWCFlatFile(lstProdPWC);
                }
                catch (Exception ex)
                {
                    Logger.LogWrite("EXC", ex);
                    Logger.LogWrite("MSG", "Return");
                    return;
                }

                //Combine Lists
                lstProdPWCAll.Concat(lstProdPWC);
            }
            #endregion

            #region Prod All
            /*
             * Get Production for each PWC all on one line for each work day
             */

            #endregion

            // Testing
            Console.WriteLine("Press key to exit");
            Console.ReadKey();
        }
    }
}
