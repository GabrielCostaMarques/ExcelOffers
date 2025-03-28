using System;
using System.Data;
using ExcelOffers.Entities;
using ExcelOffers.Factory;
using ExcelOffers.Filters;
using OfficeOpenXml;


namespace ExcelOffers
{

    class Program
    {
        public static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(@"C:\Users\gmarques\Downloads\Bloqueios R11 - V5.xlsx")))
            {
                try
                {
                    List<Product> tariff = new();

                    Console.Write("How Many offers: ");
                    int qtdOffers = int.Parse(Console.ReadLine());

                    var sheet = package.Workbook.Worksheets[0];
                    int rowCount = sheet.Dimension.Rows;


                   
                    for (int i = 2; i <= rowCount; i++)
                    {
                        tariff.Add(ProductFactory.CreateProductFromRow(sheet, tariff, i));
                    }


                    var resumoPrecos = tariff
                    .GroupBy(p => new { p.ShipName, p.Localization.EmbarkDate });



                    // 1. Agrupa por navio e data
                    // 2. Em cada grupo, ordena por Preço e pega o menor
                    // 3. (Opcional) Ordena o resultado, por exemplo, por Preço crescente
                    // 4. Pega a quantidade de ofertas que o usuário digitou
                    var cheapestByShipAndDate = tariff
                        .GroupBy(p => new { p.ShipName, p.Localization.EmbarkDate })
                        .Select(
                                g => g.OrderByDescending(p => p.Pricing.Discount)
                                .ThenBy(p=>p.Pricing.TotalFarePerPax)
                                .First()
                              )                           
                        .Take(qtdOffers)                               
                        .ToList();

                    List<Product> listFiltered = new List<Product>();


                    int row = 2;
                    foreach (var item in cheapestByShipAndDate)
                    {
                            sheet.Cells[row, 1].Value = item.ShipName;
                            sheet.Cells[row, 2].Value = item.ProductName;
                            sheet.Cells[row, 3].Value = item.CabinCategory;
                            sheet.Cells[row, 4].Value = item.CabinClass;
                            sheet.Cells[row, 5].Value = item.Pricing.Discount;
                            sheet.Cells[row, 6].Value = item.Localization.EmbarkDate.ToString("dd/MM/yyyy");
                            sheet.Cells[row, 7].Value = item.Pricing.FromToValue;

                        //colocar o restos das colunas
                            row++;
                        listFiltered.Add(item);
                    }

                    var sheetFiltered = package.Workbook.Worksheets.Add("FilteredData");



                }
                catch (Exception e)
                {
                    Console.WriteLine("Não foi possivel ler o arquivo " + e.Message);
                }

            }
        }
    }


}


