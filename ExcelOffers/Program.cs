using System;
using System.Data;
using ExcelOffers.Entities;
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

                    var sheet = packeage.Workbook.Worksheets[0];

                    int rowCount = sheet.Dimension.Rows;
                    int columnCount = sheet.Dimension.Columns;




                    for (int i = 2; i < rowCount; i++)
                    {
                        int aux = 0;

                        string shipName = sheet.Cells[i, 1].Value?.ToString() ?? "Sem dados";
                        string productName = sheet.Cells[i, 2].Value?.ToString() ?? "Sem dados";
                        string cabinCategory = sheet.Cells[i, 6].Value?.ToString() ?? "Sem dados";
                        string cabinClass = sheet.Cells[i, 7].Value?.ToString() ?? "Sem dados";

                        string portEmbarkDesembark = sheet.Cells[i, 3].Value?.ToString() ?? "Sem dados";
                        string destinyProduct = sheet.Cells[i, 4].Value?.ToString() ?? "Sem dados";
                        DateTime embarkDate = DateTime.Parse(sheet.Cells[i, 5].Value?.ToString() ?? "Sem dados");

                        decimal fromToValue = decimal.Parse(sheet.Cells[i, 8].Value?.ToString() ?? "Sem dados");
                        decimal fareSalePerPax = decimal.Parse(sheet.Cells[i, 9].Value?.ToString() ?? "Sem dados");
                        double discount = double.Parse(sheet.Cells[i, 10].Value?.ToString() ?? "Sem dados");
                        decimal nCCFPerPax = decimal.Parse(sheet.Cells[i, 11].Value?.ToString() ?? "Sem dados");
                        decimal portFarePerPax = decimal.Parse(sheet.Cells[i, 12].Value?.ToString() ?? "Sem dados");
                        decimal totalFarePerPax = decimal.Parse(sheet.Cells[i, 13].Value?.ToString() ?? "Sem dados");


                        tariff.Add(
                            new Product(
                                shipName,
                                productName,
                                cabinCategory,
                                cabinClass,
                                new Fares(
                                    fareSalePerPax,
                                    nCCFPerPax,
                                    portFarePerPax,
                                    totalFarePerPax,
                                    fromToValue,
                                    discount
                                ),
                                new Localization(
                                    portEmbarkDesembark,
                                    destinyProduct,
                                    embarkDate
                                )
                            )
                        );
                    }

                    foreach (var item in tariff)
                    {
                        Console.WriteLine(item);     
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


