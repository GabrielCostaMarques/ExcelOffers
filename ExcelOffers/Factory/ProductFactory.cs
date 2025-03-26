using ExcelOffers.Entities;
using ExcelOffers.Domain;
using OfficeOpenXml;

namespace ExcelOffers.Factory
{
    internal class ProductFactory
    {
        public static Product CreateProductFromRow(ExcelWorksheet sheet, List<Product> products, int row)
        {
            int col = 1;

            string shipName = CellValue.GetCellValue(sheet, row, ref col);
            string productName = CellValue.GetCellValue(sheet, row, ref col);
            string portEmbarkDesembark = CellValue.GetCellValue(sheet, row, ref col);
            string destinyProduct = CellValue.GetCellValue(sheet, row, ref col);
            string embarkDateStr = CellValue.GetCellValue(sheet, row, ref col);
            string cabinCategory = CellValue.GetCellValue(sheet, row, ref col);
            string cabinClass = CellValue.GetCellValue(sheet, row, ref col);

            string fromToValueStr = CellValue.GetCellValue(sheet, row, ref col);
            string fareSalePerPaxStr = CellValue.GetCellValue(sheet, row, ref col);
            string discountStr = CellValue.GetCellValue(sheet, row, ref col);
            string nCCFPerPaxStr = CellValue.GetCellValue(sheet, row, ref col);
            string portFarePerPaxStr = CellValue.GetCellValue(sheet, row, ref col);
            string totalFarePerPaxStr = CellValue.GetCellValue(sheet, row, ref col);

            DateTime embarkDate = DateTime.Parse(embarkDateStr);
            decimal fromToValue = decimal.Parse(fromToValueStr);
            decimal fareSalePerPax = decimal.Parse(fareSalePerPaxStr);
            double discount = double.Parse(discountStr);
            decimal nCCFPerPax = decimal.Parse(nCCFPerPaxStr);
            decimal portFarePerPax = decimal.Parse(portFarePerPaxStr);
            decimal totalFarePerPax = decimal.Parse(totalFarePerPaxStr);

            Product product = new Product(
                shipName,
                productName,
                cabinCategory,
                cabinClass,
                new Fares(
                    fareSalePerPax,
                    nCCFPerPax,
                    portFarePerPax,
                    totalFarePerPax,
                    fromToValue,
                    discount
                ),
                new Localization(
                    portEmbarkDesembark,
                    destinyProduct,
                    embarkDate
                )
            );

            products.Add(product);

            return product;


        }
    }
}
