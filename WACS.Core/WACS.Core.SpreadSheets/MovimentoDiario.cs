using System.Globalization;
using ClosedXML.Excel;
using Nager.Holiday;
using WACS.Entities;

namespace WACS.Core.SpreadSheets {
    class MovimentoDiario: SpreadSheet {
        
        public async void Run(string root, string path) { 
            if (path.Contains("MOVIMENTO DIÁRIO")) {
                using var workbook = new XLWorkbook(path);
                var worksheet = workbook.Worksheets.Worksheet("MOVIMENTO DIÁRIO");
                foreach (var pic in worksheet.Pictures)
                {
                    var sizes = new Dictionary<string, int[]>() {
                        {"C2", new int[] { 148, 36 }},
                        {"F2", new int[] { 360, 35 }},
                        {"P2", new int[] { 155, 33 }}
                    };
                    var size = sizes[pic.TopLeftCell + ""];
                    worksheet.Picture(pic.Name).WithSize(size[0], size[1]);
                }
                var month = DateTime.Now.ToString("MMMM", CultureInfo.CreateSpecificCulture("pt-BR")).ToUpper();

                Reset(worksheet, "A8:B38");
                Reset(worksheet, "Y8:Y38");
                foreach (var day in await Days(worksheet))
                {
                    foreach (var cell in day)
                    {
                        cell.Set();
                    }
                }
                var totals = worksheet.Range("Z41:AG41").CellsUsed().Select(item => item.Value).ToList();
                foreach (var (cell, i) in worksheet.Range("Z7:AG7").CellsUsed().Select((value, i) => (value, i)))
                {
                    cell.Value = totals[i];
                }
                worksheet.Cell("G5").Value = month;
                worksheet.Range("Z8:AG40").Value = "";
                worksheet.Range("D8:F39").Value = "";
                worksheet.Range("H8:W39").Value = "";
                worksheet.Range("AA45:AJ51").Value = "";
                worksheet.Range("Z68:AL68").Value = "";
                workbook.SaveAs(string.Format(@"{0}\MOVIMENTO DIÁRIO {1} - {2}.xlsx", root, month, DateTime.Now.Year));
                File.Delete(path); 
            }
        }

        private void Reset(IXLWorksheet? worksheet, string cell) {
            worksheet!.Range(cell).Value = "";
            worksheet!.Range(cell).Style.Fill.BackgroundColor = XLWorkbook.DefaultStyle.Fill.BackgroundColor;
            worksheet!.Range(cell).Style.Font.FontColor =  XLWorkbook.DefaultStyle.Font.FontColor;
        }

        private async Task<List<CustomCell>[]> Days(IXLWorksheet? worksheet) { 
            var cellsA = new List<CustomCell>();
            var cellsB = new List<CustomCell>();
            var cellsY = new List<CustomCell>();
            var today = DateTime.Now;
            int maxDays = DateTime.DaysInMonth(today.Year, today.Month);
            var days = new Dictionary<string, string[]>();
            using var holidayClient = new HolidayClient();
            var holidays = await holidayClient.GetHolidaysAsync(today.Year, "br");
            for (var day = 1; day <= maxDays; day++) {
                var date = new DateTime(today.Year, today.Month, day);
                var holiday = holidays!.Any(item => item.Date == date && (item.Counties == null || !item.Counties!.Any(el => el.Contains("BR-SP"))));
                var strDay = date.ToString("dddd", CultureInfo.CreateSpecificCulture("pt-BR"));
                var background = XLWorkbook.DefaultStyle.Fill.BackgroundColor;
                var color = XLWorkbook.DefaultStyle.Font.FontColor;
                if (strDay == "domingo" || holiday) {
                    background = XLColor.FromArgb(0xF2F2F2);
                    color = XLColor.Red;
                }

                cellsA.Add(new CustomCell(worksheet!, $"A{7 + day}", day, background, color));
                cellsB.Add(new CustomCell(worksheet!, $"B{7 + day}", strDay[0].ToString().ToUpper(), background, color));
                cellsY.Add(new CustomCell(worksheet!, $"Y{7 + day}", day, background, color));
            }  
            return new List<CustomCell>[] { cellsA, cellsB, cellsY };
        }
    }
}