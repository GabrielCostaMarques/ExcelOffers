using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelOffers.Entities
{
    internal class Fares
    {
        public double FareSalePerPax { get; set; }
        public double NCCFPerPax { get; set; }
        public double PortFarePerPax { get; set; }
        public double TotalFarePerPax { get; set; }
        public double FromToValue { get; set; }
        public double Discount { get; set; }


        public Fares() { }
        public Fares(double fareSalePerPax, double nCCFPerPax, double portFarePerPax, double totalFarePerPax, double fromToValue, double discount)
        {
            FareSalePerPax = fareSalePerPax;
            NCCFPerPax = nCCFPerPax;
            PortFarePerPax = portFarePerPax;
            TotalFarePerPax = totalFarePerPax;
            FromToValue = fromToValue;
            Discount = discount;
        }

        public double DownPaymentB2B()
        {
            double downPayment = FareSalePerPax * 0.25;
            return downPayment;
        }
        public double InstallmentsB2B()
        {
            double installment =( TotalFarePerPax - DownPaymentB2B())/10;
            return installment;
        }
        
        public double InstallmentsB2C()
        {
            double installment = TotalFarePerPax/10;
            return installment;
        }
    }
}
