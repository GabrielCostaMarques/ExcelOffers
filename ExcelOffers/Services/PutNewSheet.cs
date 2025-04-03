using ExcelOffers.Entities;
using OfficeOpenXml;

namespace ExcelOffers.Services
{
    internal class PutNewSheet
    {
        public void NewSheetOffer(ExcelPackage packeage, List<Product> listFiltered)
        {
            try
            {
                var newSheet = packeage.Workbook.Worksheets.Add($"FilteredData - {DateTime.Now.ToString("dd/MM/yyyy")}");
                int row = 2;
                foreach (var item in listFiltered)
                {
                    newSheet.Cells[row, 1].Value = item.ShipName;
                    newSheet.Cells[row, 2].Value = item.ProductName;
                    newSheet.Cells[row, 3].Value = item.Localization.PortEmbarkDesembark;
                    newSheet.Cells[row, 4].Value = item.Localization.DestinyProduct;
                    newSheet.Cells[row, 5].Value = item.Localization.EmbarkDate.ToString("dd/MM/yyyy");
                    newSheet.Cells[row, 6].Value = item.CabinCategory;
                    newSheet.Cells[row, 7].Value = item.CabinClass;
                    newSheet.Cells[row, 8].Value = item.Fares.FromToValue;
                    newSheet.Cells[row, 9].Value = item.Fares.FareSalePerPax;
                    newSheet.Cells[row, 10].Value = item.Fares.Discount;
                    newSheet.Cells[row, 11].Value = item.Fares.NCCFPerPax;
                    newSheet.Cells[row, 12].Value = item.Fares.PortFarePerPax;
                    newSheet.Cells[row, 13].Value = item.Fares.TotalFarePerPax;

                    newSheet.Cells[row, 8].Style.Numberformat.Format = "#,##0";
                    newSheet.Cells[row, 9].Style.Numberformat.Format = "#,##0";
                    newSheet.Cells[row, 10].Style.Numberformat.Format = "0";
                    newSheet.Cells[row, 11].Style.Numberformat.Format = "#,##0";
                    newSheet.Cells[row, 12].Style.Numberformat.Format = "#,##0";
                    newSheet.Cells[row, 13].Style.Numberformat.Format = "#,##0";

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
