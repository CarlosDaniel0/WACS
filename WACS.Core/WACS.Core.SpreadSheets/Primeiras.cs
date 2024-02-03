using System.Globalization;
using ClosedXML.Excel;

namespace WACS.Core.SpreadSheets {
    class Primeiras: SpreadSheet {
        public void Run(string root, string path) {
            if (path.Contains("PRIMEIRAS") || path.Contains("SEGUNDAS")) {
                using var workbook = new XLWorkbook(path);
                IXLWorksheet worksheet;
                if (path.Contains("PRIMEIRAS"))
                    worksheet = workbook.Worksheets.Worksheet("ENTRADAS");
                else 
                    worksheet = workbook.Worksheets.Worksheet("Plan1");

                foreach (var pic in worksheet.Pictures)
                {
                    if (path.Contains("PRIMEIRAS"))
                        worksheet.Picture(pic.Name).WithSize(118, 29);
                    if (path.Contains("SEGUNDAS"))
                        worksheet.Picture(pic.Name).WithSize(155, 38);
                }

                var today = DateTime.Now;
                var month = today.ToString("MMMM", CultureInfo.CreateSpecificCulture("pt-BR")).ToUpper();
                if (path.Contains("PRIMEIRAS")) {
                    worksheet.Cell("H3").Value = $"MÊS: {month} {today.Year}";
                    worksheet.Range("A5:J22").Value = ""; 
                }
                else { 
                    worksheet.Cell("E5").Value = $"MÊS: {month} {today.Year}";
                    worksheet.Range("B7:J21").Value = "";  
                }
                
                workbook.Save(); 
            }
        }
    }
}