using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using NumberingFormat = DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat;
using OrientationValues = DocumentFormat.OpenXml.Drawing.Charts.OrientationValues;

namespace OpenXmlAutomation
{
    public class XlBarChart(XlSheet s)
    {
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
            ChartText chartText = new ();
            RichText richText = new ();
            A.BodyProperties bodyProperties = new ();
            A.ListStyle listStyle = new ();

            A.Paragraph paragraph = new ();
            A.ParagraphProperties paragraphProperties = new ();
            A.DefaultRunProperties defaultRunProperties = new ();
            paragraphProperties.Append (defaultRunProperties);

            A.Run run = new  ();
            A.RunProperties runProperties = new() { Language = "en-GB" };
            A.Text text = new() { Text = ChartTitle };

            run.Append(runProperties);
            run.Append(text);

            paragraph.Append(paragraphProperties);
            paragraph.Append (run);

            richText.Append(bodyProperties);
            richText.Append(listStyle);
            richText.Append(paragraph);

            chartText.Append(richText);
            Overlay overlay = new () { Val = false };
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

            DrawingsPart drawingsPart = sheet.part.AddNewPart<DrawingsPart>();
            Drawing drawing = new()
            { 
                Id = sheet.part.GetIdOfPart(drawingsPart) 
            };
            sheet.part.Worksheet.Append(drawing);
            sheet.part.Worksheet.Save();

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

            // Provide the title flr the chart

            chart.Append(GenerateTitleObject());
            AutoTitleDeleted titleDeleted = new() { Val = false };
            chart.Append(titleDeleted);

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

            DataLabels dataLabels = new ();
            ShowLegendKey showLegendKey = new () { Val = false };
            ShowValue showValue = new () { Val = false };
            ShowCategoryName showCategoryName = new () { Val = false };
            ShowSeriesName showSeriesName = new () { Val = false };
            ShowPercent showPercent = new () { Val = false };
            ShowBubbleSize showBubbleSize = new () { Val = false };
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
            Delete valDelete = new() {  Val = false };
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
            chart.Append(plotArea);
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
}
