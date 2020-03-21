using System;
using System.Collections.Generic;
using System.Text;

namespace PRO_Metrics_NET.Models
{
    public class ReportsModel
    {
        public int rptID { get; set; }
        public string name { get; set; }
        public string temppath { get; set; }
        public string rootpath { get; set; }
        public string filename { get; set; }
        public string fullpath { get; set; }
        public string frq { get; set; }
        public string typ { get; set; }
    }

    public class EmployeesReportsModel
    {
        public int rptID { get; set; }
        public string email { get; set; }
        public string name { get; set; }
        public string temppath { get; set; }
        public string rootpath { get; set; }
        public string filename { get; set; }
        public string fullpath { get; set; }
        public string frq { get; set; }
        public string typ { get; set; }
    }

    public class BookingsModel
    {
        public int workDy { get; set; }
        public DateTime bookDt { get; set; }
        public int bookDly { get; set; }
        public int bookAve { get; set; }
    }

    public class SalesModel
    {
        public int workDy { get; set; }
        public DateTime SalesDt { get; set; }
        public int SalesDly { get; set; }
        public int SalesAve { get; set; }
        public int TollDly { get; set; }
        public int TollAve { get; set; }
    }

    public class ProdSumModel
    {
        public int workDy { get; set; }
        public DateTime prodDt { get; set; }
        public int jobsDly { get; set; }
        public int jobsAve { get; set; }
        public int lbsDly { get; set; }
        public int lbsAve { get; set; }
    }

    public class ProdPWCModel
    {
        public int workDy { get; set; }
        public int prodDt { get; set; }
        public string pwc { get; set; }
        public int jobs { get; set; }
        public int lbs { get; set; }
        public int brks { get; set; }
        public int setUps { get; set; }
        public int cuts { get; set; }
    }

    public class ProdALLModel
    {
        public int workDy { get; set; }
        public int prodDt { get; set; }
        public int lbs60S { get; set; }
        public int jobs60S { get; set; }
        public int lbs72S { get; set; }
        public int jobs72S { get; set; }
        public int lbsCTL { get; set; }
        public int jobsCTL { get; set; }
        public int lbsMSB { get; set; }
        public int jobsMSB { get; set; }

    }

    public class JobDetailModel
    {       
        public string pwc { get; set; }
        public DateTime Dt { get; set; }
        public int job { get; set; }
        public int lbs { get; set; }
        public int ft { get; set; }
        public int brks { get; set; }
        public int setUps { get; set; }
        public int cuts { get; set; }
        public decimal arbIn { get; set; }
        public string frm { get; set; }
        public string grd { get; set; }
        public string fnsh { get; set; }
        public decimal gauge { get; set; }
        public decimal wdth { get; set; }
    }
}
