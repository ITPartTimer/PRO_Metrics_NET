using System;
using System.IO;
using System.Reflection;
using System.Configuration;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PRO_Metrics_NET.Models;

namespace PRO_Metrics_NET.DataAccess
{
    class Flat
    {
        /*
         * Pass a List<Objects> where the list contains objects from
         * the Models class.  Use reflection to find properties and
         * values of each model and write to a flat file
         */
        public static void WriteBookingsFlatFile(List<BookingsModel> lstObj)
        {
            try
            {
                using (StreamWriter writer = new StreamWriter(ConfigurationManager.AppSettings.Get("FlatPath"), true))
                {
                    writer.WriteLine("---------- Bookings ----------");

                    foreach(Object obj in lstObj)
                    {
                        PropertyInfo[] props = obj.GetType().GetProperties();
                        foreach (PropertyInfo p in props)
                        {
                            writer.Write(p.Name + " = " + p.GetValue(obj, null).ToString() + ", ");
                        }
                        writer.WriteLine();
                    } 
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static void WriteSalesFlatFile(List<SalesModel> lstObj)
        {
            try
            {
                using (StreamWriter writer = new StreamWriter(ConfigurationManager.AppSettings.Get("FlatPath"), true))
                {
                    writer.WriteLine("---------- Sales ----------");

                    foreach (Object obj in lstObj)
                    {
                        PropertyInfo[] props = obj.GetType().GetProperties();
                        foreach (PropertyInfo p in props)
                        {
                            writer.Write(p.Name + " = " + p.GetValue(obj, null).ToString() + ", ");
                        }
                        writer.WriteLine();
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static void WriteProdSumFlatFile(List<ProdSumModel> lstObj)
        {
            try
            {
                using (StreamWriter writer = new StreamWriter(ConfigurationManager.AppSettings.Get("FlatPath"), true))
                {
                    writer.WriteLine("---------- Prod ----------");

                    foreach (Object obj in lstObj)
                    {
                        PropertyInfo[] props = obj.GetType().GetProperties();
                        foreach (PropertyInfo p in props)
                        {
                            writer.Write(p.Name + " = " + p.GetValue(obj, null).ToString() + ", ");
                        }
                        writer.WriteLine();
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static void WriteProdPWCFlatFile(List<ProdPWCModel> lstObj)
        {
            try
            {
                using (StreamWriter writer = new StreamWriter(ConfigurationManager.AppSettings.Get("FlatPath"), true))
                {
                    writer.WriteLine("---------- PWC ----------");

                    foreach (Object obj in lstObj)
                    {
                        PropertyInfo[] props = obj.GetType().GetProperties();
                        foreach (PropertyInfo p in props)
                        {
                            writer.Write(p.Name + " = " + p.GetValue(obj, null).ToString() + ", ");
                        }
                        writer.WriteLine();
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}
