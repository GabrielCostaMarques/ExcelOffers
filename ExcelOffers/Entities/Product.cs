using System;
using System.Collections.Generic;
using System.Globalization;
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
        public Fares Fares { get; set; }
        public Localization Localization { get; set; }

        public Product(string shipName, string productName, string cabinCategory, string cabinClass, Fares fares, Localization localization)
        {
            ShipName = shipName;
            ProductName = productName;
            CabinCategory = cabinCategory;
            CabinClass = cabinClass;
            Fares = fares;
            Localization = localization;
        }


        //public override string ToString()
        //{
        //    return ShipName
        //        + " - "
        //        + ProductName
        //        + " - "
        //        + CabinClass
        //        + " - "
        //        + Fares.Discount.ToString("F0", CultureInfo.InvariantCulture)
        //        + " - "
        //        + Fares.FromToValue.ToString("F2", CultureInfo.InvariantCulture)
        //        + " - "
        //        + Fares.TotalFarePerPax.ToString("F2", CultureInfo.InvariantCulture);
        //}
    }
}
