using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PRO_Metrics_NET;
using PRO_Metrics_NET.Models;

namespace PRO_Metrics_NET.DataAccess
{
    public class SQLData : Helpers
    {
        public List<EmployeesReportsModel> Get_Emp_Reports(string brh, string name)
        {
            List<EmployeesReportsModel> lst = new List<EmployeesReportsModel>();

            SqlCommand cmd = new SqlCommand();
            SqlDataReader rdr = default(SqlDataReader);

            SqlConnection conn = new SqlConnection(STRATIXDataConnString);

            using (conn)
            {
                conn.Open();

                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = "RPT_LKU_proc_Emails_ByBrh_ByRptName";
                cmd.Connection = conn;

                AddParamToSQLCmd(cmd, "@brh", SqlDbType.VarChar, 2, ParameterDirection.Input, brh);
                AddParamToSQLCmd(cmd, "@name", SqlDbType.VarChar, 25, ParameterDirection.Input, name);

                rdr = cmd.ExecuteReader();

                using (rdr)
                {
                    while (rdr.Read())
                    {
                        EmployeesReportsModel r = new EmployeesReportsModel();

                        r.rptID = (int)rdr["RptID"];
                        r.email = (string)rdr["EmpEmail"];
                        r.temppath = (string)rdr["RptTempPath"];
                        r.rootpath = (string)rdr["RptRootPath"];
                        r.filename = (string)rdr["RptFileName"];
                        r.fullpath = (string)rdr["RptFullPath"];                       

                        lst.Add(r);
                    }
                }
            }

            return lst;
        }

        public List<string> Get_PWC_ByBrh(string brh)
        {           
            List<string> pwcList = new List<string>();

            SqlCommand cmd = new SqlCommand();
            SqlDataReader rdr = default(SqlDataReader);

            SqlConnection conn = new SqlConnection(STRATIXDataConnString);

            try
            {
                using (conn)
                {
                    conn.Open();

                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "PWC_LKU_proc_Active_ByBrh";
                    cmd.Connection = conn;

                    AddParamToSQLCmd(cmd, "@brh", SqlDbType.VarChar, 3, ParameterDirection.Input, brh);

                    rdr = cmd.ExecuteReader();

                    using (rdr)
                    {
                        while (rdr.Read())
                        {
                            string s;

                            s = (string)rdr["PWC"];

                            pwcList.Add(s);
                        }
                    }
                }
            }
            catch(Exception)
            {
                throw;
            }
            finally
            {
                conn.Close();
                conn.Dispose();
            }

            return pwcList;
        }

        public List<BookingsModel> Get_Bookings_ByBrh(string days, string brh, string date1, string date2)
        {
            List<BookingsModel> lstBookings = new List<BookingsModel>();

            string cmdTxt = "";
            
            SqlCommand cmd = new SqlCommand();  
            SqlDataReader rdr = default(SqlDataReader);

            SqlConnection conn = new SqlConnection(STRATIXDataConnString);

            try
            {
                using (conn)
                {
                    conn.Open();

                    if(days == "WD")
                        cmdTxt = "SCORE_LKU_proc_Bookings_DateRange_WD_ByBrh_Agg";
                    else
                        cmdTxt = "SCORE_LKU_proc_Bookings_DateRange_All_ByBrh_Agg";

                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = cmdTxt;
                    cmd.Connection = conn;

                    AddParamToSQLCmd(cmd, "@date1", SqlDbType.DateTime, 4, ParameterDirection.Input, date1);
                    AddParamToSQLCmd(cmd, "@date2", SqlDbType.DateTime, 4, ParameterDirection.Input, date2);
                    AddParamToSQLCmd(cmd, "@brh", SqlDbType.VarChar, 3, ParameterDirection.Input, brh);

                    rdr = cmd.ExecuteReader();

                    using (rdr)
                    {
                        while(rdr.Read())
                        {
                            BookingsModel m = new BookingsModel();

                            m.workDy = (int)rdr["WORK_DY"];
                            m.bookDt = (DateTime)rdr["BOOK_DT"];
                            m.bookDly = (int)rdr["BOOK_DLY"];
                            m.bookAve = (int)rdr["BOOK_AVE"];

                            lstBookings.Add(m);
                        }
                        
                    }
                }
            }
            catch(Exception)
            {
                throw;
            }
            finally
            {
                conn.Close();
                conn.Dispose();
            }

            return lstBookings;
        }

        public List<SalesModel> Get_Sales_ByBrh(string days, string brh, string date1, string date2)
        {
            List<SalesModel> lstSales = new List<SalesModel>();

            string cmdTxt = "";

            SqlConnection conn = new SqlConnection(STRATIXDataConnString);
            SqlCommand cmd = new SqlCommand();

            SqlDataReader rdr = default(SqlDataReader);

            try
            {
                using (conn)
                {
                    conn.Open();

                    if (days == "WD")
                        cmdTxt = "SCORE_LKU_proc_Sales_DateRange_WD_ByBrh_Agg";
                    else
                        cmdTxt = "SCORE_LKU_proc_Sales_DateRange_All_ByBrh_Agg";

                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = cmdTxt;
                    cmd.Connection = conn;

                    AddParamToSQLCmd(cmd, "@date1", SqlDbType.DateTime, 4, ParameterDirection.Input, date1);
                    AddParamToSQLCmd(cmd, "@date2", SqlDbType.DateTime, 4, ParameterDirection.Input, date2);
                    AddParamToSQLCmd(cmd, "@brh", SqlDbType.VarChar, 3, ParameterDirection.Input, brh);

                    rdr = cmd.ExecuteReader();

                    using (rdr)
                    {
                        while(rdr.Read())
                        {
                            SalesModel m = new SalesModel();

                            m.workDy = (int)rdr["WORK_DY"];
                            m.SalesDt = (DateTime)rdr["SALES_DT"];
                            m.SalesDly = (int)rdr["SALES_DLY"];
                            m.SalesAve = (int)rdr["SALES_AVE"];
                            m.TollDly = (int)rdr["TOLL_DLY"];
                            m.TollAve = (int)rdr["TOLL_AVE"];

                            lstSales.Add(m);
                        }
                        
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                conn.Close();
                conn.Dispose();
            }

            return lstSales;
        }

        public List<ProdSumModel> Get_Prod_Sum(string days, string date1, string date2)
        {
            List<ProdSumModel> lstProdSum = new List<ProdSumModel>();

            string cmdTxt = "";

            SqlConnection conn = new SqlConnection(STRATIXDataConnString);
            SqlCommand cmd = new SqlCommand();

            SqlDataReader rdr = default(SqlDataReader);

            try
            {
                using (conn)
                {
                    conn.Open();

                    if (days == "WD")
                        cmdTxt = "ST_PROD_LKU_proc_DateRange_WD_Agg";
                    else
                        cmdTxt = "ST_PROD_LKU_proc_DateRange_All_Agg";

                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = cmdTxt;
                    cmd.Connection = conn;

                    AddParamToSQLCmd(cmd, "@date1", SqlDbType.DateTime, 4, ParameterDirection.Input, date1);
                    AddParamToSQLCmd(cmd, "@date2", SqlDbType.DateTime, 4, ParameterDirection.Input, date2);

                    rdr = cmd.ExecuteReader();

                    using (rdr)
                    {
                        while (rdr.Read())
                        {
                            ProdSumModel m = new ProdSumModel();

                            m.workDy = (int)rdr["WORK_DY"];
                            m.prodDt = (DateTime)rdr["PROD_DT"];
                            m.jobsDly = (int)rdr["JOBS_DLY"];
                            m.jobsAve = (int)rdr["JOBS_AVE"];
                            m.lbsDly = (int)rdr["LBS_DLY"];
                            m.lbsAve = (int)rdr["LBS_AVE"];

                            lstProdSum.Add(m);
                        }
                        
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                conn.Close();
                conn.Dispose();
            }

            return lstProdSum;
        }

        public List<ProdPWCModel> Get_Prod_PWC_ByPWC(string days, string pwc, string date1, string date2)
        {
            List<ProdPWCModel> lstProd = new List<ProdPWCModel>();

            string cmdTxt = "";

            SqlConnection conn = new SqlConnection(STRATIXDataConnString);
            SqlCommand cmd = new SqlCommand();

            SqlDataReader rdr = default(SqlDataReader);

            try
            {
                using (conn)
                {
                    conn.Open();

                    if (days == "WD")
                        cmdTxt = "ST_PROD_LKU_proc_DateRange_WD_ByPWC_Agg";
                    else
                        cmdTxt = "ST_PROD_LKU_proc_DateRange_All_ByPWC_Agg";

                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = cmdTxt;
                    cmd.Connection = conn;

                    AddParamToSQLCmd(cmd, "@date1", SqlDbType.DateTime, 4, ParameterDirection.Input, date1);
                    AddParamToSQLCmd(cmd, "@date2", SqlDbType.DateTime, 4, ParameterDirection.Input, date2);
                    AddParamToSQLCmd(cmd, "@pwc", SqlDbType.VarChar, 3, ParameterDirection.Input, pwc);

                    rdr = cmd.ExecuteReader();

                    using (rdr)
                    {
                        while (rdr.Read())
                        {
                            ProdPWCModel m = new ProdPWCModel();

                            m.workDy = (int)rdr["WORK_DY"];
                            m.prodDt = (DateTime)rdr["PROD_DT"];
                            m.pwc = rdr["PWC"].ToString();
                            m.jobs = (int)rdr["JOBS"];
                            m.lbs = (int)rdr["LBS"];
                            m.brks = (int)rdr["BRKS"];
                            m.setUps = (int)rdr["SETUPS"];
                            m.cuts = (int)rdr["CUTS"];

                            lstProd.Add(m);
                        }
                       
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                conn.Close();
                conn.Dispose();
            }

            return lstProd;
        }      

        public List<ProdALLModel> Get_Prod_All(string days, string pwc, string date1, string date2)
        {
            /*
             * pwc = ALL.  Not in PWC table in SQL Server.
             * Creates a table with a column for each PWC
             */
            List<ProdALLModel> lstProd = new List<ProdALLModel>();

            string cmdTxt = "";

            SqlConnection conn = new SqlConnection(STRATIXDataConnString);
            SqlCommand cmd = new SqlCommand();

            SqlDataReader rdr = default(SqlDataReader);

            try
            {
                using (conn)
                {
                    conn.Open();

                    if (days == "WD")
                        cmdTxt = "ST_PROD_LKU_proc_DateRange_WD_ByPWC_Agg";
                    else
                        cmdTxt = "ST_PROD_LKU_proc_DateRange_All_ByPWC_Agg";

                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = cmdTxt;
                    cmd.Connection = conn;

                    AddParamToSQLCmd(cmd, "@date1", SqlDbType.DateTime, 4, ParameterDirection.Input, date1);
                    AddParamToSQLCmd(cmd, "@date2", SqlDbType.DateTime, 4, ParameterDirection.Input, date2);
                    AddParamToSQLCmd(cmd, "@pwc", SqlDbType.VarChar, 3, ParameterDirection.Input, pwc);

                    rdr = cmd.ExecuteReader();

                    using (rdr)
                    {
                        while (rdr.Read())
                        {
                            ProdALLModel m = new ProdALLModel();

                            m.workDy = (int)rdr["WORK_DY"];
                            m.prodDt = (DateTime)rdr["PROD_DT"];
                            m.jobs60S = (int)rdr["JOBS_60S"];
                            m.lbs60S = (int)rdr["LBS_60S"];
                            m.jobs72S = (int)rdr["JOBS_72S"];
                            m.lbs72S = (int)rdr["LBS_72S"];
                            m.jobsCTL = (int)rdr["JOBS_CTL"];
                            m.lbsCTL = (int)rdr["LBS_CTL"];
                            m.jobsMSB = (int)rdr["JOBS_MSB"];
                            m.lbsMSB = (int)rdr["LBS_MSB"];

                            lstProd.Add(m);
                        }

                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                conn.Close();
                conn.Dispose();
            }

            return lstProd;
        }

        public List<CombinedModel> Get_Combined(string date1, string date2, string brh, bool toll)
        {
            List<CombinedModel> lstComb = new List<CombinedModel>();

            SqlConnection conn = new SqlConnection(STRATIXDataConnString);
            SqlCommand cmd = new SqlCommand();

            SqlDataReader rdr = default(SqlDataReader);

            try
            {
                using (conn)
                {
                    conn.Open();

                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "ST_COMBINED_LKU_proc_All_DateRange_WD_ByBrh_Agg";
                    cmd.Connection = conn;

                    AddParamToSQLCmd(cmd, "@date1", SqlDbType.DateTime, 4, ParameterDirection.Input, date1);
                    AddParamToSQLCmd(cmd, "@date2", SqlDbType.DateTime, 4, ParameterDirection.Input, date2);
                    AddParamToSQLCmd(cmd, "@brh", SqlDbType.VarChar, 3, ParameterDirection.Input, brh);

                    // Convert bool type to equivalent bit value for SQL
                    if (toll)
                        AddParamToSQLCmd(cmd, "@toll", SqlDbType.Bit, 1, ParameterDirection.Input, 1);
                    else
                        AddParamToSQLCmd(cmd, "@toll", SqlDbType.Bit, 1, ParameterDirection.Input, 0);

                    rdr = cmd.ExecuteReader();

                    using (rdr)
                    {
                        while (rdr.Read())
                        {
                            CombinedModel m = new CombinedModel();

                            m.workDy = (int)rdr["WORK_DY"];
                            m.weekDt = (DateTime)rdr["WEEK_DT"];
                            m.bkDly = (int)rdr["BOOK"];
                            m.prodDly = (int)rdr["PROD"];
                            m.slsDly = (int)rdr["SALES"];

                            lstComb.Add(m);
                        }

                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                conn.Close();
                conn.Dispose();
            }

            return lstComb;
        }

    }
}
