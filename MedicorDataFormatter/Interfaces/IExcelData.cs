using OfficeOpenXml;

namespace MedicorDataFormatter.Interfaces
{
    /// <summary>
    /// IExcelData is responsible for setting up the
    /// excel data sheets. Packages and worksheets.
    /// </summary>
    public interface IExcelData
    {
        /// <summary>
        /// Worksheet to use
        /// </summary>
        ExcelWorksheet Worksheet { get; set; }

        /// <summary>
        /// The package the worksheet exists in
        /// </summary>
        ExcelPackage Package { get; set; }
    }
}