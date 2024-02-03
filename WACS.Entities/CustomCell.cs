using ClosedXML.Excel;

namespace WACS.Entities {
    class CustomCell {
        private IXLWorksheet Worksheet;
        public string Cell;
        public dynamic Value;
        public XLColor Background;
        public XLColor Color;

        public CustomCell(IXLWorksheet worksheet, string cell, dynamic value, XLColor background, XLColor color) {
            Worksheet = worksheet;
            Cell = cell;
            Value = value;
            Background = background;
            Color = color;
        }
        
        public void Set() {
            Worksheet.Cell(Cell).Value = Value;
            Worksheet.Cell(Cell).Style.Fill.BackgroundColor = Background;
            Worksheet.Cell(Cell).Style.Font.FontColor = Color;
        }
    }
}