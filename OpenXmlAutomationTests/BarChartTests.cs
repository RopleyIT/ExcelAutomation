using OpenXmlAutomation;
namespace OpenXmlAutomationTests;

[TestClass]
public class BarChartTests
{
    [TestInitialize]
    public void Setup()
    {
        using XlDocument? doc = XlDocument.Create("C:\\tmp\\WithChart.xlsx", true);
        if (doc == null)
            Assert.Fail("XlDocument.Create failed to create file");
        XlSheet sheet1 = doc.Sheets.First();
        sheet1.SetCell("A1", "Chart title");
        sheet1.SetCell("A2", "Apples");
        sheet1.SetCell("A3", "Pears");
        sheet1.SetCell("A4", "Plums");
        sheet1.SetCell("B1", "2024");
        sheet1.SetCell("C1", "2025");
        sheet1.SetCell("B2", 0.78);
        sheet1.SetCell("B3", 0.56);
        sheet1.SetCell("B4", 0.33);
        sheet1.SetCell("C2", 0.31);
        sheet1.SetCell("C3", 0.49);
        sheet1.SetCell("C4", 0.72);
    }

    [TestMethod]
    public void CanSetupChartData()
    {
        using XlDocument doc2 = XlDocument.Open("C:\\tmp\\WithChart.xlsx");
        XlSheet sheet2 = doc2.Sheets.First();
        XlCell? b2Cell = sheet2.FindCell("B2");
        Assert.IsNotNull(b2Cell);
        Assert.IsTrue(b2Cell.IsDouble);
        Assert.AreEqual(0.78, b2Cell.AsDouble);
    }

    [TestMethod]
    public void CanCreateClusteredBarChart()
    {
        using (XlDocument doc2 = XlDocument.Open("C:\\tmp\\WithChart.xlsx"))
        {
            XlSheet sheet2 = doc2.Sheets.First();
            XlBarChart barChart = new(sheet2)
            {
                ChartTitle = "A1",
                Grouping = "clustered",
                RoundCorners = true,
                Direction = "col",
                CellArea = "E3:L20"
            };
            barChart.SeriesList.Add(new XlBarChartSeries(sheet2)
            {
                CategoryCellRange = "Sheet1!$A$2:$A$4",
                SeriesTitleCell = "Sheet1!B1",
                ValueCellRange = "Sheet1!B2:B4",
                ValueFormat = "##.#%"
            });
            barChart.SeriesList.Add(new XlBarChartSeries(sheet2)
            {
                CategoryCellRange = "Sheet1!$A$2:$A$4",
                SeriesTitleCell = "Sheet1!C1",
                ValueCellRange = "Sheet1!C2:C4",
                ValueFormat = "##.#%"
            });
            barChart.Generate();
        }
        // TODO: Create some real assertions that check for
        // the presence of the chart data.
        Assert.IsTrue(File.Exists("C:\\tmp\\WithChart.xlsx"));
    }

    [TestMethod]
    public void CanCreateMultipleBarCharts()
    {
        using XlDocument doc2 = XlDocument.Open("C:\\tmp\\WithChart.xlsx");
        XlSheet sheet2 = doc2.Sheets.First();
        XlBarChart barChart = new(sheet2)
        {
            ChartIndex = 1,
            ChartTitle = "A1",
            Grouping = "clustered",
            RoundCorners = true,
            Direction = "col",
            CellArea = "E3:L20"
        };
        barChart.SeriesList.Add(new XlBarChartSeries(sheet2)
        {
            CategoryCellRange = "Sheet1!$A$2:$A$4",
            SeriesTitleCell = "Sheet1!B1",
            ValueCellRange = "Sheet1!B2:B4",
            ValueFormat = "##.#%"
        });
        barChart.SeriesList.Add(new XlBarChartSeries(sheet2)
        {
            CategoryCellRange = "Sheet1!$A$2:$A$4",
            SeriesTitleCell = "Sheet1!C1",
            ValueCellRange = "Sheet1!C2:C4",
            ValueFormat = "##.#%"
        });
        barChart.Generate();

        // Set additional source cells
        sheet2.SetCell("A6", "2nd Chart title");
        sheet2.SetCell("A7", "Rice");
        sheet2.SetCell("A8", "Bread");
        sheet2.SetCell("A9", "Tatties");
        sheet2.SetCell("B6", "Summer");
        sheet2.SetCell("C6", "Winter");
        sheet2.SetCell("B7", 0.54);
        sheet2.SetCell("B8", 0.56);
        sheet2.SetCell("B9", 0.77);
        sheet2.SetCell("C7", 0.2);
        sheet2.SetCell("C8", 0.39);
        sheet2.SetCell("C9", 0.92);

        // Create the second bar chart
        barChart = new(sheet2)
        {
            ChartIndex = 2,
            ChartTitle = "A6",
            Grouping = "clustered",
            RoundCorners = true,
            Direction = "col",
            CellArea = "N3:U20"
        };
        barChart.SeriesList.Add(new XlBarChartSeries(sheet2)
        {
            CategoryCellRange = "Sheet1!$A$7:$A$9",
            SeriesTitleCell = "Sheet1!B6",
            ValueCellRange = "Sheet1!B7:B9",
            ValueFormat = "##.#%"
        });
        barChart.SeriesList.Add(new XlBarChartSeries(sheet2)
        {
            CategoryCellRange = "Sheet1!$A$7:$A$9",
            SeriesTitleCell = "Sheet1!C6",
            ValueCellRange = "Sheet1!C7:C9",
            ValueFormat = "##.#%"
        });
        barChart.Generate();

        // TODO: Create some real assertions that check for
        // the presence of the chart data.
        Assert.IsTrue(File.Exists("C:\\tmp\\WithChart.xlsx"));
    }
}
