using OfficeOpenXml;

namespace ExcelOffers.Domain
{
    internal class CellValue
    {
        public static string GetCellValue(ExcelWorksheet sheet, int row, ref int column)
        {
            if (true)
            {
                
            }
            string value = sheet.Cells[row, column].Value?.ToString() ?? "Sem dados";
            column++;

            return value;
        }
    }
}
