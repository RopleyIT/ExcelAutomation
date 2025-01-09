using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Office2013.Word;
using DocumentFormat.OpenXml.Office2016.Excel;
using DocumentFormat.OpenXml.Office2021.Excel.RichDataWebImage;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.RegularExpressions;
using A = DocumentFormat.OpenXml.Drawing;

namespace OpenXmlAutomation
{
    public class XlBarChart
    {
        private XlSheet sheet;
        private Chart chart;

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
        public XlBarChart(XlSheet s, string direction, string grouping, string language = "en-US")
        {
            sheet = s;

            // Add a drawing to the worksheet

            DrawingsPart drawingsPart = sheet.part.AddNewPart<DrawingsPart>();
            Drawing drawing = new()
            { 
                Id = sheet.part.GetIdOfPart(drawingsPart) 
            };
            sheet.part.Worksheet.Append(drawing);
            sheet.part.Worksheet.Save();

            // Add a chart to the drawings part and set its language

            chart = new();
            EditingLanguage editingLanguage = new()
            {
                Val = new StringValue(language)
            };
            ChartPart chartPart = drawingsPart.AddNewPart<ChartPart>();
            chartPart.ChartSpace = new ChartSpace();
            chartPart.ChartSpace.Append(editingLanguage);
            chartPart.ChartSpace.AppendChild<Chart>(chart);

            // Set up the plotting area for the bar chart

            PlotArea plotArea = chart.AppendChild<PlotArea>(new PlotArea());
            Layout layout = plotArea.AppendChild<Layout>(new Layout());
            BarChart barChart = new BarChart(barDirection(direction), barGrouping(grouping));
            plotArea.AppendChild<BarChart>(barChart);
        }

        private BarDirection barDirection(string direction)
            => new()
            {
                Val = new EnumValue<BarDirectionValues>(direction switch
                {
                    "col" => BarDirectionValues.Column,
                    "bar" => BarDirectionValues.Bar,
                    _ => throw new ArgumentException("Bar direction must be \"col\" or \"bar\"")
                })
            };

        private BarGrouping barGrouping(string grouping)
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
