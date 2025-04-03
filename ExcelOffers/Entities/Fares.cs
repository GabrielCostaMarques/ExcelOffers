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
    }
}
