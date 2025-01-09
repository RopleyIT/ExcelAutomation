using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Office2013.Word;
using DocumentFormat.OpenXml.Office2016.Excel;
using DocumentFormat.OpenXml.Office2021.Excel.RichDataWebImage;
using DocumentFormat.OpenXml.Packaging;
using System.Text.RegularExpressions;
using A = DocumentFormat.OpenXml.Drawing;

namespace OpenXmlAutomation
{
    public class XlBarChartSeries
    {
        /// <summary>
        /// The Excel cell reference for the cell containing
        /// the caption or legend for this series. Example: "Sheet1!$A$3"
        /// for cell A3 on sheet Sheet1.
        /// </summary>
        
        public string SeriesTitleCell { get; set; } = string.Empty;

        //internal BarChartSeries Generate(uint index)
        //{
        //    BarChartSeries barChartSeries = new ();
        //    DocumentFormat.OpenXml.Drawing.Charts.Index index1 = new() 
        //    { 
        //        Val = (UInt32Value)index 
        //    };
        //    Order order = new Order()
        //    {
        //        Val = (UInt32Value)index
        //    };
        //    SeriesText seriesText = new ();
        //    StringReference stringReference = new ();
        //    Formula formula = new Formula()
        //    {
        //        Text = SeriesTitleCell
        //    };
        //}
    }
}
