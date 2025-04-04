using ExcelOffers.Entities;
using OfficeOpenXml;
using ExcelOffers.Services;
using ExcelOffers.Domain;

namespace ExcelOffers.Services
{
    internal class PutNewSheet
    {
        public void NewSheetOffer(ExcelPackage packeage, List<Product> listFiltered)
        {

            CellValue cellValue = new CellValue();
            try
            {
                var newSheet = packeage.Workbook.Worksheets.Add($"FilteredData - {DateTime.Now.ToString("dd/MM/yyyy")}");
                int row = 2;
                foreach (var item in listFiltered)
                {
                    cellValue.SetCellValue(newSheet, row, 1, item.ShipName);
                    cellValue.SetCellValue(newSheet, row, 2, item.ProductName);
                    cellValue.SetCellValue(newSheet, row, 3, item.Localization.PortEmbarkDesembark);
                    cellValue.SetCellValue(newSheet, row, 4, item.Localization.DestinyProduct);
                    cellValue.SetCellValue(newSheet, row, 5, item.Localization.EmbarkDate.ToString("dd/MM/yyyy"));
                    cellValue.SetCellValue(newSheet, row, 6, item.CabinCategory);
                    cellValue.SetCellValue(newSheet, row, 7, item.CabinClass);
                    cellValue.SetCellValue(newSheet, row, 8, item.Fares.FromToValue, "#,##0");
                    cellValue.SetCellValue(newSheet, row, 9, item.Fares.FareSalePerPax, "#,##0");
                    cellValue.SetCellValue(newSheet, row, 10, item.Fares.Discount, "0");
                    cellValue.SetCellValue(newSheet, row, 11, item.Fares.NCCFPerPax, "#,##0");
                    cellValue.SetCellValue(newSheet, row, 12, item.Fares.PortFarePerPax, "#,##0");
                    cellValue.SetCellValue(newSheet, row, 13, item.Fares.TotalFarePerPax, "#,##0");
                    cellValue.SetCellValue(newSheet, row, 14, item.Fares.DownPaymentB2B(), "#,##0.00");
                    cellValue.SetCellValue(newSheet, row, 15, item.Fares.InstallmentsB2B(), "#,##0.00");
                    cellValue.SetCellValue(newSheet, row, 16, item.Fares.InstallmentsB2C(), "#,##0.00");

                    row++;
                }
                newSheet.Cells[newSheet.Dimension.Address].AutoFitColumns();


                packeage.Save();

            }
            catch (Exception e)
            {
                Console.WriteLine("Erro ao adicionar na tabela");
                Console.WriteLine(e.Message);
                Console.WriteLine(e.StackTrace);
            }
        }
    }
}
