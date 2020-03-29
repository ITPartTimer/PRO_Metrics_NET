using System;
using System.IO;
using System.Net.Mail;
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
            Logger.LogWrite("MSG", "Start: " + DateTime.Now.ToString());

            // Args initialization
            string date1 = "";
            string date2 = "";
            string brh = "";
            string days = "";
            bool emailIt = false;
            bool toll = false;

            /*
             * Copy empty Metric File from templates to metric folder
             * Create a full path to pass to Excel methods
             */
            string fileName = ConfigurationManager.AppSettings.Get("MetricsFileName");
            string templatePath = ConfigurationManager.AppSettings.Get("TemplatePath");
            string destPath = ConfigurationManager.AppSettings.Get("DestPath");

            File.Copy(Path.Combine(templatePath, fileName), Path.Combine(destPath, fileName), true);

            // Pass to Excel methods to be used in OleDb connection string
            string fullPath = Path.Combine(destPath, fileName);

            #region Args
            /*
             * arg options:
             * 1. Branch Email Days Toll
             * 2. Branch Email Days Toll StartDate StopDate
             * Branch = SW, CS, AR or MS
             * Email = true or false
             * Days = WD, ALL
             * Toll = true or false
             */
            try
            {
                if ((args.Length < 4))
                {                   
                    Logger.LogWrite("MSG", "Invalid number of args[]");
                    Logger.LogWrite("MSG", "Return on args[]");
                    return;
                }
                else if (args.Length == 4)
                {
                    // 1st arg should be the Brh
                    brh = args[0].ToString();
                    emailIt = Convert.ToBoolean(args[1]);
                    days = args[2].ToString();
                    toll = Convert.ToBoolean(args[3]);

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
                Logger.LogWrite("MSG", "Return on Args");
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

            //SQLData objSQL_PWC = new SQLData();

            List<string> lstPWC = new List<string>();

            try
            {
                lstPWC = objSQL.Get_PWC_ByBrh(brh);              
            }
            catch(Exception ex)
            {
                Logger.LogWrite("EXC", ex);
                Logger.LogWrite("MSG", "Return on PWC");
                return;
            }
            #endregion

            #region Bookings
            SQLData objSQL_Bookings = new SQLData();

            List<BookingsModel> lstBookings = new List<BookingsModel>();

            try
            {
                lstBookings = objSQL.Get_Bookings_ByBrh(days, brh, date1, date2);
                // testing
                Flat.WriteBookingsFlatFile(lstBookings);
            }
            catch (Exception ex)
            {
                Logger.LogWrite("EXC", ex);
                Logger.LogWrite("MSG", "Return on Bookings");
                return;
            }
            #endregion

            #region Sales
            List<SalesModel> lstSales = new List<SalesModel>();

            try
            {
                lstSales = objSQL.Get_Sales_ByBrh(days, brh, date1, date2);
                // testing
                Flat.WriteSalesFlatFile(lstSales);
            }
            catch (Exception ex)
            {
                Logger.LogWrite("EXC", ex);
                Logger.LogWrite("MSG", "Return on Sales");
                return;
            }
            #endregion

            #region Prod Sum
            List<ProdSumModel> lstProdSum = new List<ProdSumModel>();

            try
            {
                lstProdSum = objSQL.Get_Prod_Sum(days, date1, date2);
                // testing
                Flat.WriteProdSumFlatFile(lstProdSum);
            }
            catch (Exception ex)
            {
                Logger.LogWrite("EXC", ex);
                Logger.LogWrite("MSG", "Return on Prod Sum");
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
                    lstProdPWC = objSQL.Get_Prod_PWC_ByPWC(days, pwc, date1, date2);
                    Flat.WriteProdPWCFlatFile(lstProdPWC);

                }
                catch (Exception ex)
                {
                    Logger.LogWrite("EXC", ex);
                    Logger.LogWrite("MSG", "Return on Prod " + pwc);
                    return;
                }
               
                //Combine Lists.  Sort out PWC with LINQ, later.
                lstProdPWCAll.AddRange(lstProdPWC);
            }

            // testing
            // Should contain a list for each PWC
            Flat.WriteProdPWCFlatFile(lstProdPWCAll);
            #endregion

            #region Prod Detail
            /*
             * Get Production for each PWC all on one line for each work day
             */
            List<ProdALLModel> lstProdAll = new List<ProdALLModel>();

            try
            {
                lstProdAll = objSQL.Get_Prod_All(days, "ALL", date1, date2);
                // testing
                Flat.WriteProdSumFlatFile(lstProdSum);
            }
            catch (Exception ex)
            {
                Logger.LogWrite("EXC", ex);
                Logger.LogWrite("MSG", "Return on Prod All");
                return;
            }
            #endregion

            #region Combined
            /*
             * Get Book, Prod and Sales for all WD in date range in one resultset
             * PROD tables are brh specific, but SCORE tables require brh value
             * Report can be run with or without Toll Sales
             */
            List<CombinedModel> lstComb = new List<CombinedModel>();

            try
            {
                lstComb = objSQL.Get_Combined(date1, date2, brh, toll);
            }
            catch (Exception ex)
            {
                Logger.LogWrite("EXC", ex);
                Logger.LogWrite("MSG", "Return on Combined");
                return;
            }
            #endregion

            #region Excel          
            /*
             * Export each metric List<object> to correct XLS tab
             */
            ExcelExport objXLS = new ExcelExport();

            // Book - Tab
            try
            {
                objXLS.WriteBookings(lstBookings, fullPath);
            }
            catch (Exception ex)
            {
                Logger.LogWrite("EXC", ex);
                Logger.LogWrite("MSG", "Return on Bookings XLS");
                return;
            }

            // Sales - Tab
            try
            {
                objXLS.WriteSales(lstSales, fullPath);
            }
            catch (Exception ex)
            {
                Logger.LogWrite("EXC", ex);
                Logger.LogWrite("MSG", "Return on Sales XLS");
                return;
            }

            // Prod - Tab
            try
            {
                objXLS.WriteProdSum(lstProdSum, fullPath);
            }
            catch (Exception ex)
            {
                Logger.LogWrite("EXC", ex);
                Logger.LogWrite("MSG", "Return on Prod Sum XLS");
                return;
            }

            // Details - Tab
            try
            {
                objXLS.WriteProdAll(lstProdAll, fullPath);
            }
            catch (Exception ex)
            {
                Logger.LogWrite("EXC", ex);
                Logger.LogWrite("MSG", "Return on Prod All XLS");
                return;
            }

            /*
             * PWC Tabs
             * Iterate through lstPWC and put data in same named Tab in Excel
             */
            try
            {
                foreach (string pwc in lstPWC)
                {
                    // ALL does not have a Tab
                    if(pwc != "ALL")
                    {
                        // Using LINQ, select by pwc and order by work day
                        var queryPWC = from p in lstProdPWCAll
                                        where p.pwc == pwc
                                        orderby p.workDy ascending
                                        select p;
                        
                        // LINQ results are IEnumerable<ProdPWCModel>.  Need to cast to List<ProdPWCModel>
                        objXLS.WriteProdPWC(queryPWC.ToList(), fullPath, pwc);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogWrite("EXC", ex);
                Logger.LogWrite("MSG", "Return on Prod by PWC XLS");
                return;
            }

            // Combined - Tab
            try
            {
                objXLS.WriteCombined(lstComb, fullPath);
            }
            catch (Exception ex)
            {
                Logger.LogWrite("EXC", ex);
                Logger.LogWrite("MSG", "Return on Combined XLS");
                return;
            }
            #endregion

            #region Email
            /*
             * Emailing the file is optoinal
             * emailIt = true or false
             */
            if (emailIt)
            {
                List<EmployeesReportsModel> lstEmpReports = new List<EmployeesReportsModel>();

                SQLData objSQL_Rpts = new SQLData();

                try
                {
                    lstEmpReports = objSQL_Rpts.Get_Emp_Reports(brh, "MetricsDaily");
                }
                catch (Exception ex)
                {
                    Logger.LogWrite("EXC", ex);
                    Logger.LogWrite("MSG", "Return on Emp Rpts");
                    return;
                }

                try
                {
                    MailMessage mail = new MailMessage();

                    SmtpClient SmtpServer = new SmtpClient("smtp.office365.com");

                    mail.From = new MailAddress("sclemons@calstripsteel.com");
                    mail.Subject = brh + " - Daily";
                    mail.Body = "Report attached";

                    //Build To: line from emails in list of EmployeesReportsModel
                    foreach (EmployeesReportsModel e in lstEmpReports)
                    {
                        Logger.LogWrite("MSG", "Email: " + e.email.ToString());
                        mail.To.Add(e.email.ToString());
                    }

                    // Add attachment
                    Attachment attach;
                    attach = new Attachment(fullPath);
                    mail.Attachments.Add(attach);

                    SmtpServer.Port = 587;
                    SmtpServer.Credentials = new System.Net.NetworkCredential("sclemons@calstripsteel.com", "Smet@524");
                    SmtpServer.EnableSsl = true;

                    SmtpServer.Send(mail);
                }
                catch (Exception ex)
                {
                    Logger.LogWrite("EXC", ex);
                    Logger.LogWrite("MSG", "Return on Email");
                    return;
                }
            }
            else
                Logger.LogWrite("MSG", "No email");
            #endregion

            // Made it to the end
            Logger.LogWrite("MSG", "End: " + DateTime.Now.ToString());

            // Testing
            //Console.WriteLine("Press key to exit");
            //Console.ReadKey();
        }
    }
}
