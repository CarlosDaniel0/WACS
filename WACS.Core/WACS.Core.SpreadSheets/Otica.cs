using System.Globalization;
using ClosedXML.Excel;

namespace WACS.Core.SpreadSheets {
    class Otica: SpreadSheet {
        public void Run(string root, string path) {
            if (path.Contains("DEPÓSITO ÓTICA")) {
                using var workbook = new XLWorkbook(path);
                var worksheet = workbook.Worksheets.Worksheet("Planilha1");
                var today = DateTime.Now;
                var month = today.ToString("MMMM", CultureInfo.CreateSpecificCulture("pt-BR")).ToUpper();
                worksheet.Cell("B2").Value = $"FILIAL:________________MONSENHOR GIL________________________ MÊS:__________________________{month} {today.Year}______________________";
                if (path.Contains('1') || path.Contains('2')) {
                    worksheet.Range("B4:H13").Value = "";
                } else if (path.Contains('3')) {
                    worksheet.Range("B4:H14").Value = "";
                } else if (path.Contains('4')) {
                    worksheet.Range("B4:H16").Value = "";
                } else if (path.Contains('5')) {
                    worksheet.Range("B4:H12").Value = "";
                } 
                workbook.Save(); 
            }
        }
    }
}