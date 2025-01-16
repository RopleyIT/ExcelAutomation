using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using NumberingFormat = DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat;
using OrientationValues = DocumentFormat.OpenXml.Drawing.Charts.OrientationValues;

namespace OpenXmlAutomation;

public class XlBarChart(XlSheet s)
{
    /// <summary>
    /// Set to a unique value for each chart added
    /// </summary>
    public uint ChartIndex { get; set; }

    /// <summary>
    /// The cells the chart will overlay. A default area
    /// in the top left corner of the sheet is set up,
    /// but this property should be set for correct 
    /// positioning.
    /// </summary>
    public string CellArea { get; set; } = "A1:G14";

    /// <summary>
    /// "col" or "bar" for vertical or horizontal bars in chart
    /// </summary>
    public string Direction { get; set; } = "col";

    /// <summary>
    /// The way in which values are grouped in the bar chart
    /// "standard"      A standard single bar chart with one series
    /// "clustered"     Clustered multiple series bar chart
    /// "stacked"       Cumulative single bars for each clustered value
    /// "percent"       Like stacked, but all bars same total height
    /// </summary>
    public string Grouping { get; set; } = "clustered";

    /// <summary>
    /// The language used for text in the chart
    /// </summary>
    public string Language { get; set; } = "en-US";

    public bool RoundCorners { get; set; } = false;
    public List<XlBarChartSeries> SeriesList { get; set; } = [];
    private readonly XlSheet sheet = s;

    /// <summary>
    /// The title at the top of the chart
    /// </summary>

    public string ChartTitle { get; set; } = string.Empty;

    /// <summary>
    /// Create the OpenXML object tree for the title
    /// </summary>
    /// <returns>The title object for incorporation 
    /// into the chart's XML file</returns>

    private Title GenerateTitleObject()
    {
        var title = new Title();
        ChartText chartText = new();
        RichText richText = new();
        A.BodyProperties bodyProperties = new();
        A.ListStyle listStyle = new();

        A.Paragraph paragraph = new();
        A.ParagraphProperties paragraphProperties = new();
        A.DefaultRunProperties defaultRunProperties = new();
        paragraphProperties.Append(defaultRunProperties);

        A.Run run = new();
        A.RunProperties runProperties = new() { Language = "en-GB" };
        XlCell titleCell = sheet.FindCell(ChartTitle);
        A.Text text = new() { Text = titleCell.Value ?? string.Empty };

        run.Append(runProperties);
        run.Append(text);

        paragraph.Append(paragraphProperties);
        paragraph.Append(run);

        richText.Append(bodyProperties);
        richText.Append(listStyle);
        richText.Append(paragraph);

        chartText.Append(richText);
        Overlay overlay = new() { Val = false };
        title.Append(chartText);
        title.Append(overlay);

        return title;
    }

    /// <summary>
    /// Generate the bar chart and add it to the sheet
    /// </summary>

    public void Generate()
    {
        // Add a drawing to the worksheet
        IEnumerable<DrawingsPart> drawingsParts
            = sheet.part.GetPartsOfType<DrawingsPart>();
        DrawingsPart? drawingsPart = drawingsParts.FirstOrDefault();
        if (drawingsPart == null)
        {
            drawingsPart = sheet.part.AddNewPart<DrawingsPart>();
            drawingsPart.WorksheetDrawing = new();
            Drawing drawing = new()
            {
                Id = sheet.part.GetIdOfPart(drawingsPart)
            };
            sheet.part.Worksheet.Append(drawing);
        }


        // Add a chart to the drawings part and set its language

        EditingLanguage editingLanguage = new()
        {
            Val = new StringValue(Language)
        };
        ChartPart chartPart = drawingsPart.AddNewPart<ChartPart>();
        chartPart.ChartSpace = new ChartSpace();
        Date1904 chartDate = new() { Val = false };
        RoundedCorners roundedCorners = new()
        {
            Val = RoundCorners
        };
        chartPart.ChartSpace.Append(chartDate);
        chartPart.ChartSpace.Append(editingLanguage);
        chartPart.ChartSpace.Append(roundedCorners);
        Chart chart = chartPart.ChartSpace.AppendChild<Chart>(new Chart());

        // Provide the title for the chart

        chart.Append(GenerateTitleObject());
        chart.AutoTitleDeleted = new() { Val = true };

        // Set up the plotting area for the bar chart

        PlotArea plotArea = chart.AppendChild<PlotArea>(new PlotArea());
        plotArea.Append(new Layout());
        BarChart barChart = new
            (BarDirFromStr(Direction), BarGrpFromStr(Grouping));
        barChart.AppendChild(new VaryColors() { Val = false });
        plotArea.AppendChild<BarChart>(barChart);

        foreach (XlBarChartSeries bcs in SeriesList)
            barChart.Append(bcs.Generate((uint)SeriesList.IndexOf(bcs)));

        // Labels

        DataLabels dataLabels = new();
        ShowLegendKey showLegendKey = new() { Val = false };
        ShowValue showValue = new() { Val = false };
        ShowCategoryName showCategoryName = new() { Val = false };
        ShowSeriesName showSeriesName = new() { Val = false };
        ShowPercent showPercent = new() { Val = false };
        ShowBubbleSize showBubbleSize = new() { Val = false };
        dataLabels.Append(showLegendKey);
        dataLabels.Append(showValue);
        dataLabels.Append(showCategoryName);
        dataLabels.Append(showSeriesName);
        dataLabels.Append(showPercent);
        dataLabels.Append(showBubbleSize);
        GapWidth gapWIdth = new() { Val = (UInt16Value)150 };
        AxisId catAxisId = new()
        { Val = (UInt32Value)sheet.document.NextId() };
        AxisId valAxisId = new()
        { Val = (UInt32Value)sheet.document.NextId() };
        barChart.Append(dataLabels);
        barChart.Append(gapWIdth);
        barChart.Append(catAxisId);
        barChart.Append(valAxisId);

        // Configure the category axis

        CategoryAxis categoryAxis = new();
        AxisId clonedCatAxisId = new()
        {
            Val = catAxisId.Val
        };
        Scaling scaling = new();
        Orientation orientation = new()
        {
            Val = OrientationValues.MinMax
        };
        scaling.Append(orientation);
        Delete delete = new() { Val = false };
        AxisPosition axisPosition = new()
        {
            Val = AxisPositionValues.Bottom
        };
        A.Charts.NumberingFormat numberingFormat = new()
        {
            FormatCode = "General",
            SourceLinked = true,
        };
        MajorTickMark majorTickMark = new()
        {
            Val = TickMarkValues.Outside
        };
        MinorTickMark minorTickMark = new()
        {
            Val = TickMarkValues.None
        };
        TickLabelPosition tickLabelPosition = new()
        {
            Val = TickLabelPositionValues.NextTo
        };
        CrossingAxis crossingAxis = new()
        {
            Val = valAxisId.Val
        };
        Crosses crosses = new()
        {
            Val = CrossesValues.AutoZero
        };
        AutoLabeled autoLabeled = new()
        {
            Val = true
        };
        LabelAlignment alignment = new()
        {
            Val = LabelAlignmentValues.Center
        };
        LabelOffset labelOffset = new()
        {
            Val = (UInt16Value)100u
        };
        NoMultiLevelLabels noMultiLevelLabels = new()
        {
            Val = false
        };

        categoryAxis.Append(clonedCatAxisId);
        categoryAxis.Append(scaling);
        categoryAxis.Append(delete);
        categoryAxis.Append(axisPosition);
        categoryAxis.Append(numberingFormat);
        categoryAxis.Append(majorTickMark);
        categoryAxis.Append(minorTickMark);
        categoryAxis.Append(tickLabelPosition);
        categoryAxis.Append(crossingAxis);
        categoryAxis.Append(crosses);
        categoryAxis.Append(autoLabeled);
        categoryAxis.Append(alignment);
        categoryAxis.Append(labelOffset);
        categoryAxis.Append(noMultiLevelLabels);

        // Now set up the value axis

        ValueAxis valAxis = new();
        AxisId clonedValAxisId = new()
        {
            Val = valAxisId.Val
        };
        Scaling valScaling = new();
        Orientation valOrientation = new()
        {
            Val = OrientationValues.MinMax
        };
        valScaling.Append(valOrientation);
        Delete valDelete = new() { Val = false };
        AxisPosition valAxPos = new()
        {
            Val = AxisPositionValues.Left
        };
        MajorGridlines valMajGridLines = new();
        NumberingFormat valNumberFormat = new()
        {
            FormatCode = "##.#%",
            SourceLinked = true
        };
        MajorTickMark valMajorTickMark = new()
        {
            Val = TickMarkValues.Outside
        };
        MinorTickMark valMinorTickMark = new()
        {
            Val = TickMarkValues.None
        };
        TickLabelPosition valTickLabelPosition = new()
        {
            Val = TickLabelPositionValues.NextTo
        };
        CrossingAxis valCrossAxis = new()
        {
            Val = catAxisId.Val
        };
        Crosses valCrosses = new()
        {
            Val = CrossesValues.AutoZero
        };
        CrossBetween valBetween = new()
        {
            Val = CrossBetweenValues.Between
        };
        valAxis.Append(clonedValAxisId);
        valAxis.Append(valScaling);
        valAxis.Append(valDelete);
        valAxis.Append(valAxPos);
        valAxis.Append(valMajGridLines);
        valAxis.Append(valNumberFormat);
        valAxis.Append(valMajorTickMark);
        valAxis.Append(valMinorTickMark);
        valAxis.Append(valTickLabelPosition);
        valAxis.Append(valCrossAxis);
        valAxis.Append(valCrosses);
        valAxis.Append(valBetween);

        plotArea.Append(categoryAxis);
        plotArea.Append(valAxis);

        // Legend on bar chart

        Legend legend = new();
        LegendPosition legendPosition = new()
        {
            Val = LegendPositionValues.Right
        };
        Overlay legendOverlay = new() { Val = false };
        legend.Append(legendPosition);
        legend.Append(legendOverlay);
        PlotVisibleOnly plotVisibleOnly = new()
        {
            Val = true
        };
        DisplayBlanksAs displayBlanksAs = new()
        {
            Val = DisplayBlanksAsValues.Gap
        };
        ShowDataLabelsOverMaximum showDataLabelsOverMaximum = new()
        {
            Val = false
        };
        chart.Append(legend);
        chart.Append(plotVisibleOnly);
        chart.Append(displayBlanksAs);
        chart.Append(showDataLabelsOverMaximum);

        // Printer settings for the chart

        PrintSettings settings = new();
        A.Charts.HeaderFooter headerFooter = new();
        A.Charts.PageMargins pageMargins = new()
        {
            Left = 0.7D,
            Right = 0.7D,
            Top = 0.75D,
            Bottom = 0.75D,
            Header = 0.3D,
            Footer = 0.3D
        };
        A.Charts.PageSetup pageSetup = new();
        settings.Append(headerFooter);
        settings.Append(pageMargins);
        settings.Append(pageSetup);
        chartPart.ChartSpace.Append(settings);
        chartPart.ChartSpace.Save();

        // Now complete the entries in the drawing part. Set
        // the position of the chart on the parent sheet using
        // a two-cell anchor.

        TwoCellAnchor twoCellAnchor = new();
        drawingsPart.WorksheetDrawing.Append(twoCellAnchor);
        XlCellRef chartArea = new(CellArea);
        A.Spreadsheet.FromMarker fm = new(
            new ColumnId(XlCellRef.Index(chartArea.Column).ToString()),
            new ColumnOffset("0"),
            new RowId(XlCellRef.Index(chartArea.Row).ToString()),
            new RowOffset("0"));
        A.Spreadsheet.ToMarker tm = new(
            new ColumnId(XlCellRef.Index(chartArea.LastColumn).ToString()),
            new ColumnOffset("0"),
            new RowId(XlCellRef.Index(chartArea.LastRow).ToString()),
            new RowOffset("0"));
        twoCellAnchor.Append(fm);
        twoCellAnchor.Append(tm);
        GraphicFrame gf = new()
        {
            Macro = string.Empty,
            NonVisualGraphicFrameProperties
                = new NonVisualGraphicFrameProperties()
                {
                    NonVisualDrawingProperties
                        = new NonVisualDrawingProperties()
                        {
                            Id = ChartIndex,
                            Name = $"Chart {ChartIndex}"
                        },
                    NonVisualGraphicFrameDrawingProperties
                        = new NonVisualGraphicFrameDrawingProperties()
                },
            Transform = new()
            {
                Offset = new()
                {
                    X = 0,
                    Y = 0
                },
                Extents = new()
                {
                    Cx = 0,
                    Cy = 0
                },
            },
            Graphic = new()
            {
                GraphicData = new()
                {
                    Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart"
                }
            }
        };
        gf.Graphic.GraphicData.AppendChild(new ChartReference()
        {
            Id = drawingsPart.GetIdOfPart(chartPart)
        });
        twoCellAnchor.Append(gf);
        twoCellAnchor.Append(new A.Spreadsheet.ClientData());

        // Now save all the parts away

        sheet.part.Worksheet.Save();
        drawingsPart.WorksheetDrawing.Save();
        sheet.document.SaveWorkbook();
    }

    private static BarDirection BarDirFromStr(string direction)
        => new()
        {
            Val = new EnumValue<BarDirectionValues>(direction switch
            {
                "col" => BarDirectionValues.Column,
                "bar" => BarDirectionValues.Bar,
                _ => throw new ArgumentException("Bar direction must be \"col\" or \"bar\"")
            })
        };

    private static BarGrouping BarGrpFromStr(string grouping)
        => new()
        {
            Val = new EnumValue<BarGroupingValues>(grouping switch
            {
                "standard" => BarGroupingValues.Standard,
                "clustered" => BarGroupingValues.Clustered,
                "stacked" => BarGroupingValues.Stacked,
                "percent" => BarGroupingValues.PercentStacked,
                _ => throw new ArgumentException
                    ("Bar grouping should be one of standard, clustered, stacked, or percent")
            })
        };
}
