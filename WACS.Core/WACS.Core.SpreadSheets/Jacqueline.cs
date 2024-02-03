 

using System.Globalization;
using ClosedXML.Excel;

namespace WACS.Core.SpreadSheets {
    class Jacqueline: SpreadSheet {
        public void Run(string root, string path) {
            
            using var workbook = new XLWorkbook(path);
            IXLWorksheet worksheet;
            if (path.Contains("NEGOCIAÇÕES"))
                worksheet = workbook.Worksheets.Worksheet("Plan1");
            else 
                worksheet = workbook.Worksheets.Worksheet("ENTRADAS");

            foreach (var pic in worksheet.Pictures)
            {
                if (path.Contains("NEGOCIAÇÕES"))
                    worksheet.Picture(pic.Name).WithSize(255, 94);
                if (path.Contains("PRIMEIRAS"))
                    worksheet.Picture(pic.Name).WithSize(118, 29);
            }

            var today = DateTime.Now;
            var month = today.ToString("MMMM", CultureInfo.CreateSpecificCulture("pt-BR")).ToUpper();
            if (path.Contains("NEGOCIAÇÕES")) {
                worksheet.Cell("D4").Value = $"REF. MÊS: {month} {today.Year}       -      1ª E 2ª PAGAS";
                worksheet.Range("B7:I27").Value = ""; 
            } else if (path.Contains("PRIMEIRAS")) {
                worksheet.Cell("A3").Value = $"RELATÓRIO DE PRIMEIRAS - JACQUELINE - {month.Substring(0, 3)} {today.Year.ToString().Substring(2, 4)}";
                worksheet.Range("B5:F31").Value = "";
            }
            workbook.Save();
        }
    }
}