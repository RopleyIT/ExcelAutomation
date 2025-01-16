using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace OpenXmlAutomation;

/// <summary>
/// Implementation of a single series of points on a bar chart.
/// There may be multiple of these under a BarChart if the bar chart
/// has clustered series, stacked or percent series grouping.
/// </summary>
public class XlBarChartSeries(XlSheet s)
{
    private readonly XlSheet sheet = s;

    /// <summary>
    /// The Excel cell reference for the cell containing
    /// the caption or legend for this series. Example: "Sheet1!$A$3"
    /// for cell A3 on sheet Sheet1.
    /// </summary>
    public string SeriesTitleCell { get; set; } = string.Empty;

    /// <summary>
    /// The vector of cells that define the Categories
    /// </summary>
    public string CategoryCellRange { get; set; } = string.Empty;

    /// <summary>
    /// The vector of cells that define the values in each category
    /// </summary>
    public string ValueCellRange { get; set; } = string.Empty;

    /// <summary>
    /// The format for the values on the bar chart, e.g. "##.#%"
    /// </summary>
    public string ValueFormat { get; set; } = string.Empty;

    internal BarChartSeries Generate(uint index)
    {
        BarChartSeries barChartSeries = new();
        DocumentFormat.OpenXml.Drawing.Charts.Index chartIndex = new()
        {
            Val = (UInt32Value)index
        };
        Order order = new()
        {
            Val = (UInt32Value)index
        };
        SeriesText seriesText = new();
        StringReference stringReference = new();
        Formula formula = new()
        {
            Text = SeriesTitleCell
        };

        // Now find the value of the title cell reference
        // so that the string cache can be built for it

        StringCache titleCache = sheet.document
            .StringCacheFromRange(SeriesTitleCell);
        stringReference.Append(formula);
        stringReference.Append(titleCache);
        seriesText.Append(stringReference);

        // Specify whether to invert the bar chart

        InvertIfNegative invert = new()
        {
            Val = false
        };

        // Populate the category axis with category titles

        CategoryAxisData catAxisData = new();
        StringReference catRef = new();
        Formula catFormula = new()
        {
            Text = CategoryCellRange
        };
        StringCache categories = sheet.document
            .StringCacheFromRange(CategoryCellRange);
        catRef.Append(catFormula);
        catRef.Append(categories);
        catAxisData.Append(catRef);

        // Now fetch the values for each category

        Values values = new();
        NumberReference numRef = new();
        Formula valFormula = new()
        {
            Text = ValueCellRange
        };
        NumberingCache numCache = sheet.document
            .NumberingCacheFromCellRange(ValueCellRange, ValueFormat);
        numRef.Append(valFormula);
        numRef.Append(numCache);
        values.Append(numRef);

        // Non OpenXML Microsoft extensions

        //BarSerExtensionList bsExtList = new ();
        //BarSerExtension bsExt = new()
        //{
        //    Uri = "{C3380CC4-5D6E-409C-BE32-E72D297353CC}"
        //};
        //bsExt.AddNamespaceDeclaration
        //    ("c16", "http://schemas/microsoft.com/office/drawing/2014/chart");
        //OpenXmlUnknownElement unkElem = new OpenXmlUnknownElement()
        //    ("<c16:uniqueId val=\"{00000002-2359-4C33-91BD-6E50FC1E9826}\" xmlns:c16=\"http://schemas.microsoft.com/office/drawing/2014/chart\" />"); 

        barChartSeries.Append(chartIndex);
        barChartSeries.Append(order);
        barChartSeries.Append(seriesText);
        barChartSeries.Append(invert);
        barChartSeries.Append(catAxisData);
        barChartSeries.Append(values);
        //barChartSeries.Append(barSerExtensionList);
        return barChartSeries;
    }
}
