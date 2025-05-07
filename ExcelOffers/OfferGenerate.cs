using ExcelOffers.ExcelOffers.Domain.Entities;
using ExcelOffers.ExcelOffers.Infrastructure;
using ExcelOffers.Factory;
using ExcelOffers.Services;
using OfficeOpenXml;

namespace ExcelOffers
{
    internal class OfferGenerate
    {
        public void GenerateOffers() {

            using (var package = new ExcelPackage(new FileInfo(@"C:\Users\gmarques\Downloads\Bloqueios\Bloqueios R11 - V12.xlsx")))
            {
                FilterFactory filter = new FilterFactory();
                List<Product> tariff = new();
                ExcelProductWriter putNewSheet = new ExcelProductWriter();

                try
                {
                    Console.Write("How Many offers: ");
                    int qtdOffers = int.Parse(Console.ReadLine());

                    var sheet = package.Workbook.Worksheets[0];
                    int rowCount = sheet.Dimension.Rows;


                    for (int i = 2; i <= rowCount; i++)
                    {
                        ProductFactory.CreateProductFromRow(sheet, tariff, i);
                    }

                    var sortedOffers = filter.FilterTariff(tariff, qtdOffers);
                    putNewSheet.NewSheetOffer(package, sortedOffers);
                }


                catch (Exception e)
                {
                    Console.WriteLine("Não foi possivel ler o arquivo " + e.Message);
                    Console.WriteLine(e.StackTrace);
                }
            }
        }
    }
}
