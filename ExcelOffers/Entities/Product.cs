using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelOffers.Entities
{
    internal class Product
    {
        public string ShipName { get; set; }
        public string ProductName { get; set; }
        public string CabinCategory { get; set; }
        public string CabinClass { get; set; }
        public Fares Pricing { get; set; }
        public Localization Localization { get; set; }

        public Product(string shipName, string productName, string cabinCategory, string cabinClass, Fares pricing, Localization localization)
        {
            ShipName = shipName;
            ProductName = productName;
            CabinCategory = cabinCategory;
            CabinClass = cabinClass;
            Pricing = pricing;
            Localization = localization;
        }


        public override string ToString()
        {
            return ShipName
                + ", "
                + ProductName
                + ", "
                + Pricing.FromToValue
                +", "
                +Pricing.TotalFarePerPax;
        }
    }
}
