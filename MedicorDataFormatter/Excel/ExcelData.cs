using MedicorDataFormatter.Interfaces;
using OfficeOpenXml;
using System;
using System.IO;
using System.Linq;

namespace MedicorDataFormatter.Excel
{
    /// <summary>
    /// Excel setup finds a workbook/package and retrieves a worksheet from it.
    /// </summary>
    public class ExcelData : IExcelData
    {
        #region Properties
        /// <summary>
        /// The excel workbook to open
        /// </summary>
        public ExcelPackage Package { get; }
        /// <summary>
        /// The worksheet to access
        /// </summary>
        public ExcelWorksheet Worksheet { get; }
        #endregion

        #region Constructors
        /// <summary>
        /// Setup the excel workbook. Find the excel file as a package
        /// and then find the excel worksheet by name.
        /// </summary>
        /// <param name="path"></param>
        /// <param name="worksheetName"></param>
        public ExcelData(string path, string worksheetName)
        {
            // check that the parameters supplied are valid. If either null or empty then error would occur stop here.
            if (string.IsNullOrWhiteSpace(path))
                throw new ArgumentNullException(nameof(path), "The path supplied is blank. Enter a valid path");

            if (string.IsNullOrWhiteSpace(worksheetName))
                throw new ArgumentNullException(nameof(worksheetName), "The sheet name supplied is blank. " +
                                                                        "Enter a valid sheet name");

            Package = new ExcelPackage(new FileInfo(path));

            // find the worksheet by name
            Worksheet = Package.Workbook.Worksheets
                .FirstOrDefault(x => x.Name.Equals(worksheetName, StringComparison.CurrentCultureIgnoreCase));

            // throw error if no worksheet has been found
            if (Worksheet == null)
                throw new FileNotFoundException("Could not find the worksheet. Check the file path and worksheet " +
                                                "name are correct and are correct.");
        }
        #endregion
    }
}
