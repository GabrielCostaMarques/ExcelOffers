using ExcelOffers.Entities;

namespace ExcelOffers.Factory
{
    internal class FilterFactory
    {
        public List<Product> FilterTariff(List<Product> list, int takeOffers)
        {
            var sortedList =
                  list.OrderByDescending(p => p.Fares.Discount)
                      .ThenBy(p => p.Fares.TotalFarePerPax)
                      .GroupBy(p => new { p.ShipName, p.Localization.EmbarkDate })
                      .Select(g => g.MinBy(p => p.Fares.TotalFarePerPax))
                      .Take(takeOffers)
                      .ToList();

            return sortedList;
        }
    }
}
