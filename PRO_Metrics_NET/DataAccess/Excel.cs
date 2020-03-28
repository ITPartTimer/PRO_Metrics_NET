using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Data.OleDb;
using System.Text;
using System.Threading.Tasks;
using PRO_Metrics_NET.Models;

namespace PRO_Metrics_NET.DataAccess
{
    public class ExcelExport
    {
        public void WriteBookings(List<BookingsModel> lstBookings, string fullPath)
        {
            // initialize text used in OleDbCommand
            string cmdText = "";

            string excelConnString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fullPath + @";Extended Properties=""Excel 8.0;HDR=YES;""";

            using (OleDbConnection eConn = new OleDbConnection(excelConnString))
            {
                try
                {
                    eConn.Open();

                    OleDbCommand eCmd = new OleDbCommand();

                    eCmd.Connection = eConn;

                    // Loop through each record and add yesterday details to XLS
                    foreach (BookingsModel m in lstBookings)
                    {
                        // Use parameters to insert into XLS
                        cmdText = "Insert into [Book$] (WORK_DY,BOOK_DT,BOOK_DLY,BOOK_AVE) Values(@dy,@dt,@bk_dly,@bk_ave)";

                        eCmd.CommandText = cmdText;

                        eCmd.Parameters.AddRange(new OleDbParameter[]
                        {
                                    new OleDbParameter("@dy", m.workDy),
                                    new OleDbParameter("@dt", m.bookDt),
                                    new OleDbParameter("@bk_dly", m.bookDly),
                                    new OleDbParameter("@bk_ave", m.bookAve),                                  
                        });

                        eCmd.ExecuteNonQuery();

                        // Need to clear Parameters on each pass
                        eCmd.Parameters.Clear();
                    }
                }
                catch (OleDbException)
                {
                    throw;
                }
                catch(Exception)
                {
                    throw;
                }
                finally
                {
                    eConn.Close();
                    eConn.Dispose();
                }
            }
        }

        public void WriteSales(List<SalesModel> lstSales, string fullPath)
        {
            // initialize text used in OleDbCommand
            string cmdText = "";

            string excelConnString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fullPath + @";Extended Properties=""Excel 8.0;HDR=YES;""";

            using (OleDbConnection eConn = new OleDbConnection(excelConnString))
            {
                try
                {
                    eConn.Open();

                    OleDbCommand eCmd = new OleDbCommand();

                    eCmd.Connection = eConn;

                    // Loop through each record and add yesterday details to XLS
                    foreach (SalesModel m in lstSales)
                    {
                        // Use parameters to insert into XLS
                        cmdText = "Insert into [Sales$] (WORK_DY,SALES_DT,SALES_DLY,SALES_AVE,TOLL_DLY,TOLL_AVE) Values(@dy,@dt,@sl_dly,@sl_ave,@tl_dly,@tl_ave)";

                        eCmd.CommandText = cmdText;

                        eCmd.Parameters.AddRange(new OleDbParameter[]
                        {
                                    new OleDbParameter("@dy", m.workDy),
                                    new OleDbParameter("@dt", m.SalesDt),
                                    new OleDbParameter("@sl_dly", m.SalesDly),
                                    new OleDbParameter("@sl_ave", m.SalesAve),
                                    new OleDbParameter("@sl_dly", m.TollDly),
                                    new OleDbParameter("@sl_ave", m.TollAve),
                        });

                        eCmd.ExecuteNonQuery();

                        // Need to clear Parameters on each pass
                        eCmd.Parameters.Clear();
                    }
                }
                catch (OleDbException)
                {
                    throw;
                }
                catch (Exception)
                {
                    throw;
                }
                finally
                {
                    eConn.Close();
                    eConn.Dispose();
                }
            }
        }

        public void WriteProdSum(List<ProdSumModel> lstProd, string fullPath)
        {
            // initialize text used in OleDbCommand
            string cmdText = "";

            string excelConnString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fullPath + @";Extended Properties=""Excel 8.0;HDR=YES;""";

            using (OleDbConnection eConn = new OleDbConnection(excelConnString))
            {
                try
                {
                    eConn.Open();

                    OleDbCommand eCmd = new OleDbCommand();

                    eCmd.Connection = eConn;

                    // Loop through each record and add yesterday details to XLS
                    foreach (ProdSumModel m in lstProd)
                    {
                        // Use parameters to insert into XLS
                        cmdText = "Insert into [Prod$] (WORK_DY,PROD_DT,JOBS_DLY,JOBS_AVE,LBS_DLY,LBS_AVE) Values(@dy,@dt,@jb_dly,@jb_ave,@lbs_dly,@lbs_ave)";

                        eCmd.CommandText = cmdText;

                        eCmd.Parameters.AddRange(new OleDbParameter[]
                        {
                                    new OleDbParameter("@dy", m.workDy),
                                    new OleDbParameter("@dt", m.prodDt),
                                    new OleDbParameter("@jb_dly", m.jobsDly),
                                    new OleDbParameter("@jb_ave", m.jobsAve),
                                    new OleDbParameter("@lbs_dly", m.lbsDly),
                                    new OleDbParameter("@lbs_ave", m.lbsAve),
                        });

                        eCmd.ExecuteNonQuery();

                        // Need to clear Parameters on each pass
                        eCmd.Parameters.Clear();
                    }
                }
                catch (OleDbException)
                {
                    throw;
                }
                catch (Exception)
                {
                    throw;
                }
                finally
                {
                    eConn.Close();
                    eConn.Dispose();
                }
            }
        }

        public void WriteProdPWC(List<ProdPWCModel> lstProd, string fullPath, string pwc)
        {
            // initialize text used in OleDbCommand
            string cmdText = "";

            string excelConnString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fullPath + @";Extended Properties=""Excel 8.0;HDR=YES;""";

            using (OleDbConnection eConn = new OleDbConnection(excelConnString))
            {
                try
                {
                    eConn.Open();

                    OleDbCommand eCmd = new OleDbCommand();

                    eCmd.Connection = eConn;

                    // Loop through each record and add yesterday details to XLS
                    foreach (ProdPWCModel m in lstProd)
                    {
                        // Use parameters to insert into XLS
                        cmdText = "Insert into [" + pwc + "$] (WORK_DY,PROD_DT,PWC,JOBS,LBS,BRKS,SETUPS,CUTS) Values(@dy,@dt,@pwc,@jb,@lbs,@brks,@setups,@cuts)";

                        eCmd.CommandText = cmdText;

                        eCmd.Parameters.AddRange(new OleDbParameter[]
                        {
                                    new OleDbParameter("@dy", m.workDy),
                                    new OleDbParameter("@dt", m.prodDt),
                                    new OleDbParameter("@pwc", m.pwc),
                                    new OleDbParameter("@jb", m.jobs),
                                    new OleDbParameter("@lbs", m.lbs),
                                    new OleDbParameter("@brks", m.brks),
                                    new OleDbParameter("@setups", m.setUps),
                                    new OleDbParameter("@cuts", m.cuts),
                        });

                        eCmd.ExecuteNonQuery();

                        // Need to clear Parameters on each pass
                        eCmd.Parameters.Clear();
                    }
                }
                catch (OleDbException)
                {
                    throw;
                }
                catch (Exception)
                {
                    throw;
                }
                finally
                {
                    eConn.Close();
                    eConn.Dispose();
                }
            }
        }

        public void WriteProdAll(List<ProdALLModel> lstProd, string fullPath)
        {
            // initialize text used in OleDbCommand
            string cmdText = "";

            string excelConnString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fullPath + @";Extended Properties=""Excel 8.0;HDR=YES;""";

            using (OleDbConnection eConn = new OleDbConnection(excelConnString))
            {
                try
                {
                    eConn.Open();

                    OleDbCommand eCmd = new OleDbCommand();

                    eCmd.Connection = eConn;

                    // Loop through each record and add yesterday details to XLS
                    foreach (ProdALLModel m in lstProd)
                    {
                        // Use parameters to insert into XLS
                        cmdText = "Insert into [Details$] (WORK_DY,PROD_DT,JOBS_60S,LBS_60S,JOBS_72S,LBS_72S,JOBS_CTL,LBS_CTL,JOBS_MSB,LBS_MSB) Values(@dy,@dt,@jb_60s,@lbs_60s,@jb_72s,@lbs_72s,@jb_ctl,@lbs_ctl,@jb_msb,@lbs_msb)";

                        eCmd.CommandText = cmdText;

                        eCmd.Parameters.AddRange(new OleDbParameter[]
                        {
                                    new OleDbParameter("@dy", m.workDy),
                                    new OleDbParameter("@dt", m.prodDt),
                                    new OleDbParameter("@jb_60s", m.jobs60S),
                                    new OleDbParameter("@lbs_60s", m.lbs60S),
                                    new OleDbParameter("@jb_72s", m.jobs72S),
                                    new OleDbParameter("@lbs_72s", m.lbs72S),
                                    new OleDbParameter("@jb_ctl", m.jobsCTL),
                                    new OleDbParameter("@lbs_ctl", m.lbsCTL),
                                    new OleDbParameter("@jb_msb", m.jobsMSB),
                                    new OleDbParameter("@lbs_msb", m.lbsMSB),
                        });

                        eCmd.ExecuteNonQuery();

                        // Need to clear Parameters on each pass
                        eCmd.Parameters.Clear();
                    }
                }
                catch (OleDbException)
                {
                    throw;
                }
                catch (Exception)
                {
                    throw;
                }
                finally
                {
                    eConn.Close();
                    eConn.Dispose();
                }
            }
        }

        public void WriteCombined(List<CombinedModel> lstComb, string fullPath)
        {
            // initialize text used in OleDbCommand
            string cmdText = "";

            string excelConnString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fullPath + @";Extended Properties=""Excel 8.0;HDR=YES;""";

            using (OleDbConnection eConn = new OleDbConnection(excelConnString))
            {
                try
                {
                    eConn.Open();

                    OleDbCommand eCmd = new OleDbCommand();

                    eCmd.Connection = eConn;

                    // Loop through each record and add yesterday details to XLS
                    foreach (CombinedModel m in lstComb)
                    {
                        // Use parameters to insert into XLS
                        cmdText = "Insert into [Combined$] (WORK_DY,WEEK_DT,BOOK,PROD,SALES) Values(@dy,@dt,@bk,@prd,@sls)";

                        eCmd.CommandText = cmdText;

                        eCmd.Parameters.AddRange(new OleDbParameter[]
                        {
                                    new OleDbParameter("@dy", m.workDy),
                                    new OleDbParameter("@dt", m.weekDt),
                                    new OleDbParameter("@bk", m.bkDly),
                                    new OleDbParameter("@prd", m.prodDly),
                                    new OleDbParameter("@sls", m.slsDly),
                        });

                        eCmd.ExecuteNonQuery();

                        // Need to clear Parameters on each pass
                        eCmd.Parameters.Clear();
                    }
                }
                catch (OleDbException)
                {
                    throw;
                }
                catch (Exception)
                {
                    throw;
                }
                finally
                {
                    eConn.Close();
                    eConn.Dispose();
                }
            }
        }

    }
    
}
