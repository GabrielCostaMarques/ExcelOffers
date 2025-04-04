using ExcelOffers.Entities;
using ExcelOffers.Factory;
using ExcelOffers.Services;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelOffers
{
    internal class OfferGenerate
    {
        public void GenerateOffers() {

            using (var packeage = new ExcelPackage(new FileInfo(@"C:\Users\gmarques\Downloads\Bloqueios\Bloqueios R11 - V5.xlsx")))
            {
                Filter filter = new Filter();
                List<Product> tariff = new();
                PutNewSheet putNewSheet = new PutNewSheet();

                try
                {

                    Console.Write("How Many offers: ");
                    int qtdOffers = int.Parse(Console.ReadLine());

                    var sheet = packeage.Workbook.Worksheets[0];
                    int rowCount = sheet.Dimension.Rows;


                    for (int i = 2; i <= rowCount; i++)
                    {
                        ProductFactory.CreateProductFromRow(sheet, tariff, i);
                    }


                    var sortedOffers = filter.FilterTariff(tariff, qtdOffers);
                    putNewSheet.NewSheetOffer(packeage, sortedOffers);


                    Console.WriteLine("---------------------------");
                    Console.WriteLine("---------------------------");
                    Console.WriteLine();
                    Console.WriteLine("Tabela realizada com Sucesso!");
                    Console.WriteLine(); ;
                    Console.WriteLine("---------------------------");
                    Console.WriteLine("---------------------------");
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
