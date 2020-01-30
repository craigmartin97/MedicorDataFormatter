using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;

namespace MedicorDataFormatter.Excel
{
    /// <summary>
    /// Excel reader accesses an excel file and opens it 
    /// The class is used to retrieve the data
    /// </summary>
    public class ExcelFormatter
    {
        #region Fields
        private readonly ExcelWorksheet worksheet;
        private readonly ExcelPackage package;
        #endregion

        #region Constructors
        /// <summary>
        /// Find the excel file and retrieve a workbook by the given name
        /// </summary>
        /// <param name="path">The path, where to find the excel file</param>
        /// <param name="worksheetName">The name of the worksheet to use</param>
        public ExcelFormatter(string path, string worksheetName)
        {
            // check that the parameters supplied are valid. If either null or empty then error would occur stop here.
            if (string.IsNullOrWhiteSpace(path))
                throw new ArgumentNullException(nameof(path), "The path supplied is blank. Enter a valid path");

            if (string.IsNullOrWhiteSpace(worksheetName))
                throw new ArgumentNullException(nameof(worksheetName), "The sheet name supplied is blank. " +
                    "Enter a valid sheet name");

            package = new ExcelPackage(new FileInfo(path));

            // find the worksheet by name
            worksheet = package.Workbook.Worksheets
                .FirstOrDefault(x => x.Name.Equals(worksheetName, StringComparison.CurrentCultureIgnoreCase));

            // throw error if no worksheet has been found
            if (worksheet == null)
                throw new FileNotFoundException("Could not find the worksheet. Check the file path and worksheet " +
                    "name are correct and are correct.");
        }
        #endregion

        #region Read Excel Data
        public void FormatExcelHealthFile()
        {
            int rows = worksheet.Dimension.Rows;
            int cols = worksheet.Dimension.Columns;

            for (int col = 1; col <= cols; col++) //each column
            {
                for (int row = 1; row <= rows; row++) // each row
                {
                    // get the value from the cell
                    object content = worksheet.Cells[row, col].Value;

                    if (content == null) // the cell has nothing in, so add the phrase needed
                    {
                        string insertValue = GetNullReferenceValue(col);
                        worksheet.Cells[row, col].Value = insertValue;
                    }
                    else
                    {
                        //check that the current data is a double, have to with EPPlus library as datetime cell returns double
                        bool isDouble = double.TryParse(content.ToString(), out double numDateTime);

                        if (isDouble)
                        {
                            DateTime currentCellDateTime = DateTime.FromOADate(numDateTime);

                            if (currentCellDateTime.Hour <= 9) // suspect 12 hour format carry on.
                            {
                                DateTime? findNext = null;
                                int nextCol = col; // nextCol, as in the prev or next col to the current
                                bool isEnd;

                                if (cols == col) // last col
                                {
                                    do
                                    {
                                        isEnd = worksheet.Dimension.Start.Column == nextCol;
                                        nextCol--;

                                        object prev = worksheet.Cells[row, nextCol].Value;
                                        if (prev == null)
                                            continue;

                                        if (double.TryParse(prev.ToString(), out double nextDateAsNum))
                                            findNext = DateTime.FromOADate(nextDateAsNum);
                                    }
                                    while (findNext == null && !isEnd);

                                    if (findNext != null && findNext.HasValue)
                                    {
                                        if (findNext.Value.Hour <= 9) // if this is also less than 9, then it must be a morning (AM)
                                            continue;

                                        // add 12 hours to make 24 hr
                                        DateTime temp = currentCellDateTime.AddHours(12);
                                        if (findNext >= temp)
                                        {
                                            worksheet.Cells[row, col].Value = temp;
                                            ApplyRedBorderStyle(row, col);
                                        }
                                    }
                                }
                                else // first or middle cols
                                {
                                    do
                                    {
                                        isEnd = cols == nextCol;
                                        nextCol++;

                                        object next = worksheet.Cells[row, nextCol].Value;
                                        if (next == null)
                                            continue;

                                        if (double.TryParse(next.ToString(), out double nextDateAsNum))
                                            findNext = DateTime.FromOADate(nextDateAsNum);
                                    }
                                    while (findNext == null && !isEnd);

                                    if (findNext != null && findNext.HasValue)
                                    {
                                        if (findNext.Value.Hour <= 9) // if this is also less than 9, then it must be a morning (AM)
                                            continue;

                                        // add 12 hours to make 24 hr
                                        DateTime temp = currentCellDateTime.AddHours(12);
                                        if (temp <= findNext)
                                        {
                                            worksheet.Cells[row, col].Value = temp;
                                            ApplyRedBorderStyle(row, col);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            package.Save(); // save the excel file.
        }
        #endregion

        #region Null String Value
        /// <summary>
        /// Returns a string reason based upon an integer
        /// </summary>
        /// <param name="col">The column number</param>
        /// <returns>Returns a string based upon an int</returns>
        private string GetNullReferenceValue(int col)
        {
            return col switch
            {
                1 => "Time into theatre",
                2 => "Time of Anaesthetic Start",
                3 => "Time into Theatre",
                4 => "Time out of Theatre",
                5 => "Time into Recovery",
                6 => "Time Out of Recovery",
                _ => "Error",
            };
        }
        #endregion

        #region Apply Stylings
        /// <summary>
        /// Apply a thick red border to a specified cell on the worksheet
        /// </summary>
        /// <param name="row">The row the cell is on</param>
        /// <param name="col">The column the cell is on</param>
        public void ApplyRedBorderStyle(int row, int col)
        {
            worksheet.Cells[row, col].Style.Border.Right.Style = ExcelBorderStyle.Thick;
            worksheet.Cells[row, col].Style.Border.Left.Style = ExcelBorderStyle.Thick;
            worksheet.Cells[row, col].Style.Border.Top.Style = ExcelBorderStyle.Thick;
            worksheet.Cells[row, col].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
            worksheet.Cells[row, col].Style.Border.Right.Color.SetColor(Color.Red);
            worksheet.Cells[row, col].Style.Border.Left.Color.SetColor(Color.Red);
            worksheet.Cells[row, col].Style.Border.Top.Color.SetColor(Color.Red);
            worksheet.Cells[row, col].Style.Border.Bottom.Color.SetColor(Color.Red);
        }
        #endregion
    }
}
