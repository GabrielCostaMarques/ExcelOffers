using System;
using System.Data;
using System.Security.Cryptography;
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



                    for (int i = 2; i <= rowCount; i++)
                    {
                        tariff.Add(ProductFactory.CreateProductFromRow(sheet, tariff, i));
                    }   


                    Console.WriteLine("Finalizou a leitura do Excel.");


                    var sortedList = tariff
                      .OrderByDescending(p => p.Fares.Discount)
                      .ThenBy(p => p.Fares.TotalFarePerPax) 
                      .GroupBy(p => new { p.ShipName, p.Localization.EmbarkDate }) 
                      .Select(g => g.First()) 
                      .Take(qtdOffers) 
                      .ToList();


                    var sheetFiltered = packeage.Workbook.Worksheets.Add($"FilteredData - {DateTime.Now.ToString("dd/MM/yyyy")}");

                    try
                    {
                        int row = 2;
                        foreach (var item in sortedList)
                        {
                            sheetFiltered.Cells["A2"].LoadFromCollection(sortedList, false);

                            sheetFiltered.Cells[row, 8].Style.Numberformat.Format = "0";
                            sheetFiltered.Cells[row, 9].Style.Numberformat.Format = "0";
                            sheetFiltered.Cells[row, 10].Style.Numberformat.Format = "0";
                            sheetFiltered.Cells[row, 11].Style.Numberformat.Format = "0";
                            sheetFiltered.Cells[row, 12].Style.Numberformat.Format = "0";
                            sheetFiltered.Cells[row, 13].Style.Numberformat.Format = "0";
                            row++;
                        }
                        sheetFiltered.Cells[sheetFiltered.Dimension.Address].AutoFitColumns();

                        packeage.Save();

                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("Erro ao adicionar na tabela");
                        Console.WriteLine(e.Message);
                        Console.WriteLine(e.StackTrace);
                    }
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


