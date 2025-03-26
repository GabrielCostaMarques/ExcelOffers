using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelOffers.Entities
{
    internal class Fares
    {
        public decimal FareSalePerPax { get; set; }
        public decimal NCCFPerPax { get; set; }
        public decimal PortFarePerPax { get; set; }
        public decimal TotalFarePerPax { get; set; }
        public decimal FromToValue { get; set; }
        public double Discount { get; set; }

        public Fares(decimal fareSalePerPax, decimal nCCFPerPax, decimal portFarePerPax, decimal totalFarePerPax, decimal fromToValue, double discount)
        {
            FareSalePerPax = fareSalePerPax;
            NCCFPerPax = nCCFPerPax;
            PortFarePerPax = portFarePerPax;
            TotalFarePerPax = totalFarePerPax;
            FromToValue = fromToValue;
            Discount = discount*100;
        }

        public Fares(decimal totalFarePerPax, decimal fromToValue, double discount)
        {
            TotalFarePerPax = totalFarePerPax;
            FromToValue = fromToValue;
            Discount = discount;
        }
    }
}
