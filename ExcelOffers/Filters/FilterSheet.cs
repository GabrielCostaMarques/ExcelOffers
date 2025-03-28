using ExcelOffers.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelOffers.Filters
{
    internal class FilterSheet
    {
        public List<Product> DiscountFilter(List<Product>list)
        {
            list.OrderByDescending(t => t.Pricing.Discount);
            return list;  
        }

        public List<Product> DateEmbarkFilter(List<Product> list)
        {
            list.OrderByDescending(t => t.Localization.EmbarkDate);
            return list;
        }
    }
}
