using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXmlAutomation;
namespace OpenXmlAutomationTests
{
    [TestClass]
    public sealed class XlDOcumentTests
    {
        [TestMethod]
        public void CanOpenSpreadSheet()
        {
            using XlDocument doc = new("C:\\tmp\\Book1.xlsx");
            Assert.IsInstanceOfType<XlDocument>(doc);
        }

        [TestMethod]
        public void EnumeratesSheetsByName()
        {
            using XlDocument doc = new("C:\\tmp\\Book1.xlsx");
            List<string> sheetNames = doc
                .Sheets
                .Select(s => s.Name)
                .ToList();
            Assert.AreEqual(1, sheetNames.Count);
            Assert.AreEqual("Sheet1", sheetNames[0]);
        }

        [TestMethod]
        public void CanOpenSheet()
        {
            using XlDocument doc = new("C:\\tmp\\Book1.xlsx");
            var sheetPart = doc
                .Sheets
                .FirstOrDefault(s => s.Name == "Sheet1");
            Assert.IsNotNull(sheetPart);
        }

        [TestMethod]
        public void CanAddSheetToWorkbook()
        {
            File.Copy("C:\\tmp\\Book1.xlsx", "C:\\tmp\\Book2.xlsx", true);
            using XlDocument doc = new("C:\\tmp\\Book2.xlsx");
            doc.AddSheet("SheetTwo");
            List<string> sheetNames = doc
                .Sheets
                .Select(s=>s.Name)
                .ToList();
            Assert.AreEqual(2, sheetNames.Count);
            Assert.AreEqual("Sheet1", sheetNames[0]);
            Assert.AreEqual("SheetTwo", sheetNames[1]);
        }

        [TestMethod]
        public void CanCreateNewSpreadsheet()
        {
            using XlDocument? doc = XlDocument.Create("C:\\tmp\\Book3.xlsx", true);
            if (doc == null)
                Assert.Fail("XlDocument.Create failed to create file");
            List<string> sheetNames = doc
                .Sheets
                .Select(s => s.Name)
                .ToList();
            Assert.AreEqual(1, sheetNames.Count);
            Assert.AreEqual("Sheet1", sheetNames[0]);
        }

        [TestMethod]
        public void CannotOverwriteSpreadsheet()
        {
            File.Copy("C:\\tmp\\Book1.xlsx", "C:\\tmp\\Book2.xlsx", true);
            using XlDocument? doc = XlDocument.Create("C:\\tmp\\Book2.xlsx", false);
            Assert.IsNull(doc);
        }

        [TestMethod]
        public void CanAddCells()
        {
            using XlDocument? doc = XlDocument.Create("C:\\tmp\\Book2.xlsx", true);
            Assert.IsNotNull(doc);
            XlSheet sheet = doc.Sheets.First(s => s.Name == "Sheet1");
            XlCell b2 = sheet.FindCell("B2");
            b2.Value = "I am cell B2";
            XlCell c3 = sheet.FindCell("C3");
            c3.Value = "I am cell C3";
            Assert.IsInstanceOfType<XlCell>(b2);
            Assert.IsInstanceOfType<XlCell>(c3);
        }

        [TestMethod]
        public void CanReadCells()
        {
            using XlDocument? doc = XlDocument.Create("C:\\tmp\\Book2.xlsx", true);
            Assert.IsNotNull(doc);
            XlSheet sheet = doc.Sheets.First(s => s.Name == "Sheet1");
            sheet.FindCell("B2").Value = "I am cell B2";
            XlCell c3 = sheet.FindCell("C3");
            c3.Value = "I am cell C3";
            Assert.AreEqual("I am cell B2", sheet.FindCell("B2").Value);
            Assert.AreEqual("I am cell C3", c3.Value);
        }

        [TestMethod]
        public void CanReopenAndReadCells()
        {
            using (XlDocument? doc = XlDocument.Create("C:\\tmp\\Book2.xlsx", true))
            {
                Assert.IsNotNull(doc);
                XlSheet sheet = doc.Sheets.First(s => s.Name == "Sheet1");
                XlCell b2 = sheet.SetCell("B2", "I am cell B2");
                XlCell c3 = sheet.SetCell("C3", "I am cell C3");
            }
            using XlDocument doc2 = new("C:\\tmp\\Book2.xlsx");
            XlSheet sheet1 = doc2.Sheets.First();
            XlCell? b2Cell = sheet1.FindCell("B2");
            XlCell? c3Cell = sheet1.FindCell("C3");
            Assert.IsNotNull(b2Cell);
            Assert.AreEqual("I am cell B2", b2Cell.Value);
            Assert.IsNotNull(c3Cell);
            Assert.AreEqual("I am cell C3", c3Cell.Value);
        }

        [TestMethod]
        public void CanAmendCells()
        {
            using XlDocument? doc = XlDocument.Create("C:\\tmp\\Book2.xlsx", true);
            Assert.IsNotNull(doc);
            XlSheet sheet = doc.Sheets.First(s => s.Name == "Sheet1");
            XlCell c3 = sheet.SetCell("C3", "I am cell C3");
            c3.Value = "I am modified";
            Assert.AreEqual("I am modified", c3.Value);
        }

        [TestMethod]
        public void CanReopenAndReadAmendedCells()
        {
            using (XlDocument? doc = XlDocument.Create("C:\\tmp\\Book2.xlsx", true))
            {
                Assert.IsNotNull(doc);
                XlSheet sheet = doc.Sheets.First(s => s.Name == "Sheet1");
                XlCell b2 = sheet.SetCell("B2", "I am cell B2");
                b2.Value = "I am modified";
            }
            using XlDocument doc2 = new("C:\\tmp\\Book2.xlsx");
            XlSheet sheet1 = doc2.Sheets.First();
            XlCell? b2Cell = sheet1.FindCell("B2");
            Assert.IsNotNull(b2Cell);
            Assert.AreEqual("I am modified", b2Cell.Value);
        }

        [DataTestMethod]
        [DataRow("A", 0)]
        [DataRow("AA", 26)]
        [DataRow("ZZ", 701)]
        public void ColumnIndicesCorrect(string cellRef, int expected)
        {
            int col = XlCell.ToColIndex(cellRef);
            Assert.AreEqual(expected, col);
        }

        [DataTestMethod]
        [DataRow("A1", 0u, 0u)]
        [DataRow("AA23", 26u, 22u)]
        [DataRow("ZZ999", 701u, 998u)]
        public void CellHasCorrectIndices(string cellRef, uint col, uint row)
        {
            using XlDocument? doc = XlDocument.Create("C:\\tmp\\Book2.xlsx", true);
            Assert.IsNotNull(doc);
            XlSheet sheet = doc.Sheets.First(s => s.Name == "Sheet1");
            XlCell b2 = sheet.FindCell(cellRef);
            Assert.AreEqual(row, b2.RowIndex);
            Assert.AreEqual(col, b2.ColumnIndex);
        }

        [TestMethod]
        public void IdenticalStringsAreShared()
        {
            using (XlDocument? doc = XlDocument.Create("C:\\tmp\\Book2.xlsx", true))
            {
                Assert.IsNotNull(doc);
                XlSheet sheet = doc.Sheets.First(s => s.Name == "Sheet1");
                XlCell b2 = sheet.SetCell("B2", "I am a cell");
                XlCell c3 = sheet.SetCell("C3", "I am a cell");
            }
            using XlDocument doc2 = new("C:\\tmp\\Book2.xlsx");
            Assert.AreEqual(1, doc2
                .SharedStringTableContents
                .Count(s => s == "I am a cell"));
        }

        [TestMethod]
        public void SharedStringsDeletedAfterFinalUse()
        {
            using (XlDocument? doc = XlDocument.Create("C:\\tmp\\Book2.xlsx", true))
            {
                Assert.IsNotNull(doc);
                XlSheet sheet = doc.Sheets.First(s => s.Name == "Sheet1");
                XlCell b2 = sheet.SetCell("B2", "I am a cell");
                XlCell c3 = sheet.SetCell("C3", "I am a cell");
            }
            using (XlDocument doc2 = new("C:\\tmp\\Book2.xlsx"))
            {
                XlSheet sheet = doc2.Sheets.First(s => s.Name == "Sheet1");
                XlCell? b2Cell = sheet.FindCell("B2");
                Assert.IsNotNull(b2Cell);
                b2Cell.Value = "I am a different cell";
                Assert.AreEqual(1, doc2
                    .SharedStringTableContents
                    .Count(s => s == "I am a cell"));
                Assert.AreEqual(1, doc2
                    .SharedStringTableContents
                    .Count(s => s == "I am a different cell"));
            }
            using (XlDocument doc3 = new("C:\\tmp\\Book2.xlsx"))
            {
                XlSheet sheet = doc3.Sheets.First(s => s.Name == "Sheet1");
                sheet.SetCell("B2", null);
                Assert.AreEqual(1, doc3
                    .SharedStringTableContents
                    .Count(s => s == "I am a cell"));
                Assert.AreEqual(0, doc3
                    .SharedStringTableContents
                    .Count(s => s == "I am a different cell"));
            }
            using (XlDocument doc4 = new("C:\\tmp\\Book2.xlsx"))
            {
                XlSheet sheet = doc4.Sheets.First(s => s.Name == "Sheet1");
                sheet.SetCell("C3", null);
                Assert.AreEqual(0, doc4
                    .SharedStringTableContents
                    .Count(s => s == "I am a cell"));
                Assert.AreEqual(0, doc4
                    .SharedStringTableContents
                    .Count(s => s == "I am a different cell"));
            }
        }
    }
}
