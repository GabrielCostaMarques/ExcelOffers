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

            using (var packeage = new ExcelPackage(new FileInfo(@"C:\Users\gmarques\Downloads\Bloqueios R11 - V5.xlsx")))
            {
                try
                {
                    List<Product> tariff = new();

                    Console.Write("How Many offers: ");
                    int qtdOffers = int.Parse(Console.ReadLine());

                    var sheet = packeage.Workbook.Worksheets[0];
                    


                    for (int i = 2; i < qtdOffers; i++)
                    {
                        tariff.Add(ProductFactory.CreateProductFromRow(sheet, tariff, i));
                    }


                    var resultFilter = tariff.OrderByDescending(t => t.Pricing.Discount).ThenBy(t => t.Localization.EmbarkDate);


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


