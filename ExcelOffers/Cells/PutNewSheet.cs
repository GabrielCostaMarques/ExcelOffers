using ExcelOffers.Entities;
using OfficeOpenXml;
using ExcelOffers.Services;
using ExcelOffers.Domain;
using OfficeOpenXml.Table;
using System;
using System.Security.Cryptography.X509Certificates;

namespace ExcelOffers.Services
{
    internal class PutNewSheet
    {
        public void NewSheetOffer(ExcelPackage packeage, List<Product> listFiltered)
        {

            CellValue cellValue = new CellValue();
            try
            {


                var newSheet = packeage.Workbook.Worksheets.Add($"FilteredData - {DateTime.Now.ToString("dd-MM-yyyy")}");


                newSheet.Cells[1, 1].Value = "Nome do Navio";
                newSheet.Cells[1, 2].Value = "Nome do Produto";
                newSheet.Cells[1, 3].Value = "Porto Embarque/ Desembarque";
                newSheet.Cells[1, 4].Value = "Tipo de Produto";
                newSheet.Cells[1, 5].Value = "Data de embarque";
                newSheet.Cells[1, 6].Value = " Categoria";
                newSheet.Cells[1, 7].Value = "Classe da Categoria";
                newSheet.Cells[1, 8].Value = "Tarifa DE por hósp.";
                newSheet.Cells[1, 9].Value = "Tarifa promocional por hósp.";
                newSheet.Cells[1, 10].Value = "Desconto %";
                newSheet.Cells[1, 11].Value = "NCCF por hósp.";
                newSheet.Cells[1, 12].Value = "Tx. Porto por hósp.";
                newSheet.Cells[1, 13].Value = "Valor total por hósp.";
                newSheet.Cells[1, 14].Value = "Entrada B2B";
                newSheet.Cells[1, 15].Value = "Parcelamento B2B";
                newSheet.Cells[1, 16].Value = "Entrada B2C";


                int row = 2;

                int column=1;
                int starRow = 1;
                int startColumn = 1;
                foreach (var item in listFiltered)
                {

                    column = 1;

                    cellValue.SetCellValue(newSheet, row, ref column, item.ShipName);
                    cellValue.SetCellValue(newSheet, row, ref column, item.ProductName);
                    cellValue.SetCellValue(newSheet, row, ref column, item.Localization.PortEmbarkDesembark);
                    cellValue.SetCellValue(newSheet, row, ref column, item.Localization.DestinyProduct);
                    cellValue.SetCellValue(newSheet, row, ref column, item.Localization.EmbarkDate, "dd/MM/yyyy");
                    cellValue.SetCellValue(newSheet, row, ref column, item.CabinCategory);
                    cellValue.SetCellValue(newSheet, row, ref column, item.CabinClass);
                    cellValue.SetCellValue(newSheet, row, ref column, item.Fares.FromToValue, "#,##0");
                    cellValue.SetCellValue(newSheet, row, ref column, item.Fares.FareSalePerPax, "#,##0.00");
                    cellValue.SetCellValue(newSheet, row, ref column, item.Fares.Discount, "0");
                    cellValue.SetCellValue(newSheet, row, ref column, item.Fares.NCCFPerPax, "#,##0.00");
                    cellValue.SetCellValue(newSheet, row, ref column, item.Fares.PortFarePerPax, "#,##0.00");
                    cellValue.SetCellValue(newSheet, row, ref column, item.Fares.TotalFarePerPax, "#,##0");
                    cellValue.SetCellValue(newSheet, row, ref column, item.Fares.DownPaymentB2B(), "#,##0.00");
                    cellValue.SetCellValue(newSheet, row, ref column, item.Fares.InstallmentsB2B(), "#,##0.00");
                    cellValue.SetCellValue(newSheet, row, ref column, item.Fares.InstallmentsB2C(), "#,##0.00");

                    row++;
                }

                newSheet.Cells[newSheet.Dimension.Address].AutoFitColumns();

                var tableRange = new ExcelAddress(starRow, startColumn, row, column-1);
                var dataRange = newSheet.Cells[tableRange.Address];

                var tabela = newSheet.Tables.Add(dataRange, "MinhaTabela");
                tabela.TableStyle = TableStyles.Medium9;
                tabela.ShowHeader = true;


                packeage.Save();

            }
            catch (Exception e)
            {
                Console.WriteLine("Erro ao adicionar na tabela");
                Console.WriteLine(e.Message);
                Console.WriteLine(e.StackTrace);
                return;
            }
        }
    }
}
