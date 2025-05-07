using ExcelOffers.ExcelOffers.Domain.Entities;

namespace ExcelOffers.Factory
{
    internal class FilterFactory
    {
        public List<Product> FilterTariff(List<Product> list, int takeOffers)
        {
            // Pega os 10 mais baratos
            var cheapestList = list
                .OrderBy(p => p.Fares.TotalFarePerPax)
                .Take(10)
                .ToList();

            // Agrupa e seleciona o melhor de cada grupo (maior desconto e menor preço)
            var groupedBestList = list
                .GroupBy(p => new { p.ShipName, p.Localization.EmbarkDate })
                .Select(g => g
                    .OrderByDescending(p => p.Fares.Discount)
                    .ThenBy(p => p.Fares.TotalFarePerPax)
                    .First()
                )
                .OrderByDescending(p => p.Fares.Discount)
                .Take(takeOffers-10)
                .ToList();

            // Junta os 50 (ou qualquer valor de takeOffers) + os 10 mais baratos
            var finalList = groupedBestList
                .Concat(cheapestList)
                .ToList();

            return finalList;
        }

    }
}
