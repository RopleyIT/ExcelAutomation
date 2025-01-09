using ClosedXML;
using ClosedXML.Excel;
namespace ExcelAutomation
{
    public class AutomationEngine
    {
        readonly string xlFile = "C:\\tmp\\book1.xlsx";

        public AutomationEngine(string file)
            => xlFile = file;

        public AutomationEngine() { }

        public void Rewrite()
        {
            using var workBook = new XLWorkbook(xlFile);
            var workSheet = workBook.Worksheet(1);
            var rnd = new Random();
            var offset = rnd.Next(12);
            for(int i = 0; i < 12; i++)
            {
                int rowNum = 1 + ((i + offset) % 12);
                var row = workSheet.Row(rowNum);
                var xCell = row.Cell(1);
                var yCell = row.Cell(2);
                xCell.Value = i / 6.0;
                yCell.Value = Math.Sin(i * Math.PI / 6.0);
            }
            workBook.Save();
        }
    }
}
