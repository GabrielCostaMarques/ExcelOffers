using ExcelOffers.Entities;
using ExcelOffers.Factory;
using ExcelOffers.Services;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Diagnostics.Metrics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelOffers
{
    internal class OfferGenerate
    {
        public void GenerateOffers() {

            using (var packeage = new ExcelPackage(new FileInfo(@"C:\Users\gmarques\Downloads\Bloqueios\Bloqueios R11 - V7.xlsx")))
            {
                FilterFactory filter = new FilterFactory();
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
