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
            DateTime embarkDate = DateTime.Parse(CellValue.GetCellValue(sheet, row, ref col));
            string cabinCategory = CellValue.GetCellValue(sheet, row, ref col);
            string cabinClass = CellValue.GetCellValue(sheet, row, ref col);



            decimal fromToValue = decimal.Parse(CellValue.GetCellValue(sheet, row, ref col));
            decimal fareSalePerPax = decimal.Parse(CellValue.GetCellValue(sheet, row, ref col));
            double discount = double.Parse(CellValue.GetCellValue(sheet, row, ref col));
            decimal nCCFPerPax = decimal.Parse(CellValue.GetCellValue(sheet, row, ref col));
            decimal portFarePerPax = decimal.Parse(CellValue.GetCellValue(sheet, row, ref col));
            decimal totalFarePerPax = decimal.Parse(CellValue.GetCellValue(sheet, row, ref col));

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
