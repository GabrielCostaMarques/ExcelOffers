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




                    string GetCellValue(ExcelWorksheet sheet, int row, int col)
                        => sheet.Cells[row, col].Value?.ToString() ?? "Sem dados";

                    for (int i = 2; i < rowCount; i++)
                    {
                        string shipName = GetCellValue(sheet, i, 1);
                        string productName = GetCellValue(sheet, i, 2);
                        string portEmbarkDesembark = GetCellValue(sheet, i, 3);
                        string destinyProduct = GetCellValue(sheet, i, 4);
                        string embarkDateStr = GetCellValue(sheet, i, 5);
                        string cabinCategory = GetCellValue(sheet, i, 6);
                        string cabinClass = GetCellValue(sheet, i, 7);
                        string fromToValueStr = GetCellValue(sheet, i, 8);
                        string fareSalePerPaxStr = GetCellValue(sheet, i, 9);
                        string discountStr = GetCellValue(sheet, i, 10);
                        string nCCFPerPaxStr = GetCellValue(sheet, i, 11);
                        string portFarePerPaxStr = GetCellValue(sheet, i, 12);
                        string totalFarePerPaxStr = GetCellValue(sheet, i, 13);

                        DateTime embarkDate = DateTime.Parse(embarkDateStr);
                        decimal fromToValue = decimal.Parse(fromToValueStr);
                        decimal fareSalePerPax = decimal.Parse(fareSalePerPaxStr);
                        double discount = double.Parse(discountStr);
                        decimal nCCFPerPax = decimal.Parse(nCCFPerPaxStr);
                        decimal portFarePerPax = decimal.Parse(portFarePerPaxStr);
                        decimal totalFarePerPax = decimal.Parse(totalFarePerPaxStr);

                        // Processa os dados conforme necessário...
                    
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


