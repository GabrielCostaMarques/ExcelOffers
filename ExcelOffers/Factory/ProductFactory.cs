using ExcelOffers.Domain;
using ExcelOffers.Entities;
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
            double fareFitPerPax = double.Parse(CellValue.GetCellValue(sheet, row, ref col));
            double fromToValue = double.Parse(CellValue.GetCellValue(sheet, row, ref col));
            double fareSalePerPax = double.Parse(CellValue.GetCellValue(sheet, row, ref col));
            double discount = double.Parse(CellValue.GetCellValue(sheet, row, ref col))*100;
            double nCCFPerPax = double.Parse(CellValue.GetCellValue(sheet, row, ref col));
            double portFarePerPax = double.Parse(CellValue.GetCellValue(sheet, row, ref col));
            double serviceFarePerPax = double.Parse(CellValue.GetCellValue(sheet, row, ref col));
            double totalFarePerPax = double.Parse(CellValue.GetCellValue(sheet, row, ref col));

            Product product = new Product(
                shipName,
                productName,
                cabinCategory,
                cabinClass,
                new Fares(
                    fareSalePerPax,
                    nCCFPerPax,
                    portFarePerPax,
                    serviceFarePerPax,
                    totalFarePerPax,
                    fareFitPerPax,
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
