﻿using OfficeOpenXml;

namespace ExcelOffers.Domain
{
    internal class CellValue
    {
        public static string GetCellValue(ExcelWorksheet sheet, int row, ref int column)
        {
            string value = sheet.Cells[row, column].Value?.ToString() ?? "Sem dados";
            column++;

            return value;
        }

        public void SetCellValue(ExcelWorksheet sheet, int row, ref int col, object value, string format = null)
        {
            var cell = sheet.Cells[row, col];

            // Força o valor como DateTime se ele for string em formato de data
            if (value is string strVal && DateTime.TryParse(strVal, out var dateVal))
            {
                cell.Value = dateVal;
            }
            else
            {
                cell.Value = value;
            }

            col++;

            cell.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

            if (!string.IsNullOrEmpty(format))
            {
                cell.Style.Numberformat.Format = format;
            }
        }

    }
}

