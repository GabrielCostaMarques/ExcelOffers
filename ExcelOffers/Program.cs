using System;
using System.Data;
using ExcelOffers.Entities;
using ExcelOffers.Factory;
using OfficeOpenXml;


namespace ExcelOffers
{

    class Program
    {
        public static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var packeage = new ExcelPackage(new FileInfo(@"C:\Users\gmarques\Downloads\Bloqueios R11 - V6.xlsx")))
            {
                try
                {
                    List<Product> tariff = new();

                    Console.Write("How Many offers: ");
                    int qtdOffers = int.Parse(Console.ReadLine());

                    var sheet = packeage.Workbook.Worksheets[0];
                    int rowCount = sheet.Dimension.Rows;

                    ProductFactory factory = new ProductFactory();


                    for (int i = 2; i <= rowCount; i++)
                    {
                        tariff.Add(ProductFactory.CreateProductFromRow(sheet, tariff, i));
                    }   


                    Console.WriteLine("Finalizou a leitura do Excel.");

                    var resultFilter = tariff
                        .GroupBy(p => new { p.ShipName, p.Localization.EmbarkDate })
                        .Select(g => g.OrderByDescending(p => p.Pricing.Discount).ThenBy(p => p.Pricing.TotalFarePerPax).First())
                        .Take(qtdOffers)
                        .ToList();


                    var take = resultFilter.Take(qtdOffers).Select(t => t);

                    foreach (var item in take)
                    {
                        Console.WriteLine(item);
                        Console.WriteLine();
                    }




                }
                catch (Exception e)
                {
                    Console.WriteLine("Não foi possivel ler o arquivo " + e.Message);
                }

            }
        }
    }


}


