# ExcelAutomation API

This library is designed to provide a less painful access to the OpenXML library
used to create or manipulate Open office documents, such as those created using
Microsoft 365 tools.

In particular, this small library focusses on APIs used to create or edit spreadsheets.
With the library it is possible to create or open existing .XLSX files, to add or alter
the worksheets within the docuemnt's top-level workbook, to manipulate
the content of cells within worksheets, and to create charts linked to cells in those
worksheets.

At the current time, the library supports just bar charts as a chart type, but may in
due course expand to include other chart types if uptake and interest exist.

## A simple example

The following sample code shows how to use the library to create a new empty
spreadsheet, to populate some grids of cells with text and with numbers as tables,
and to add bar charts to the sheet based on those tables as their data sources.

First, the new spreadsheet is created using a static method of the `XlDocument`
class. Since this is intended to be a new spreadsheet, the method is invoked with
its second argument set to true, which causes a pre-existing file with the same 
name to be overwritten if that file exists. The new spreadsheet file will contain 
a workbook with a single worksheet in it named `Sheet 1` that holds no populated
cells.

```
using OpenXmlAutomation;

namespace OpenXmlDemo;

public class Demo
{
    public void CreateDemoSpreadsheet()
    {
        using XlDocument? doc = XlDocument.Create
            ("C:\\tmp\\WithChart.xlsx", true);
      
        . . .
    }
}
```
Next we need to access the worksheet that was created automatically. The `XlDocument`
object contains a `Sheets` property that is a list of `XlSheet` objects, in this
case containing just one `XlSheet`:

```
        . . .

        using XlDocument? doc = XlDocument.Create
            ("C:\\tmp\\WithChart.xlsx", true);
        XlSheet sheet = doc.Sheets[0];
```
We shall now create a small table with text cells for row and column headers, and
number cells for some values:

```
        . . .

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
}
```
Note that the end of the function closes the `using` statement, causing the
`XlDocument` object's `Dispose()` method to be invoked. This in turn causes
the spreadsheet file to be saved to disk.

## Creating a barchart

We shall now add a bar chart to the same spreadsheet. Since the file already
exists, we shall open it rather than create it, then obtain a reference to
the same sheet containing the table we just created. IN the following code,
we assume we are inside some other method of a class in our application:

```
     using XlDocument doc2 = XlDocument.Open("C:\\tmp\\WithChart.xlsx");
     XlSheet sheet2 = doc2.Sheets[0];
```
Now we create a bar chart object and make it a child of the sheet. Notice
that we also set the cell to use when selecting a title for the bar chart,
we set the chart grouping to clustered as there are multiple rows and
columns of data, and we arrange for the bars to be vertical rather than
horizontal. We also arrange for the chart area to have rounded corners,
and we choose the range of cells that determine the area where the chart
will be drawn on the sheet:

```
    XlBarChart barChart = new(sheet2)
    {
        ChartIndex = 1,
        ChartTitle = "A1",
        Grouping = "clustered",
        RoundCorners = true,
        Direction = "col",
        CellArea = "E3:L20"
    };
```
Our table has two columns of values in columns B and C, and has the
category labels for each row alongside those values in column A. We
create a bar chart series for each column of data, since we are going
to show clustered columns on the bar chart:

```
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
```
Lastly, we invoke the `Generate` method to create the barchart within 
the spreadsheet, saving the modified contents away into the object tree:

```
    barChart.Generate();
```
As before, the closing brace of the function containing the `using`
statement that opened the spreadsheet file will save the contents of
the file to disk.

Note that if you wish to create multiple charts on the same sheet,
you can use the same code structure as above. Just remember to give
each successive bar chart an increasing value for its `ChartIndex`
property when you create the bar chart object.