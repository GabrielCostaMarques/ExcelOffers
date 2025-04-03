using System;
using System.Data;
using System.Security.Cryptography;
using ExcelOffers.Domain;
using ExcelOffers.Entities;
using ExcelOffers.Factory;
using ExcelOffers.Services;
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
                    int rowCount = sheet.Dimension.Rows;


                    for (int i = 2; i <= rowCount; i++)
                    {
                        ProductFactory.CreateProductFromRow(sheet, tariff, i); 
                    }   

                    Filter filter = new Filter();
                    var sortedOffers = filter.FilterTariff(tariff, qtdOffers);


                    PutNewSheet putNewSheet = new PutNewSheet();
                    putNewSheet.NewSheetOffer(packeage, sortedOffers);


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


