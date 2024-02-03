using System.Globalization;
using ClosedXML.Excel;

namespace WACS.Core.SpreadSheets {
    class Vendas: SpreadSheet {
        public void Run(string root, string path) {
            if (path.Contains("VENDAS FILIAL") || path.Contains("VENDAS TERESINA")) {
                using var workbook = new XLWorkbook(path);
                var worksheet = workbook.Worksheets.Worksheet("VENDAS");
                foreach (var pic in worksheet.Pictures)
                {
                    worksheet.Picture(pic.Name).WithSize(117, 34);
                }

                var today = DateTime.Now;
                var month = today.ToString("MMMM", CultureInfo.CreateSpecificCulture("pt-BR")).ToUpper();
                if (path.Contains("FILIAL")) {
                    worksheet.Cell("H4").Value = month;
                    worksheet.Range("B6:K29").Value = "";
                } else if (path.Contains("TERESINA")) {
                    worksheet.Cell("E4").Value = $"PERÍODO: {today.Year}";
                    worksheet.Range("A6:K89").Value = "";
                }
                workbook.Save(); 
            }
        }
    }
}