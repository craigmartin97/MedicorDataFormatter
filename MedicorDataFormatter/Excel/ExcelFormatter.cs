using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Drawing;
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
        /// <summary>
        /// The sheet to engage with
        /// </summary>
        private readonly ExcelWorksheet worksheet;

        /// <summary>
        /// The package, excel workbook, to work with
        /// </summary>
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
                                            ApplyRedBorderStyle(row, col, ExcelBorderStyle.Thick, Color.Red);
                                        }
                                    }
                                }
                                else // first or middle cols
                                {
                                    do
                                    {
                                        isEnd = cols == nextCol; // true
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
                                        if (findNext.Value.Hour <= 9) // if this is also less than or equal to 9, then it must be a morning (AM)
                                            continue;

                                        // add 12 hours to make 24 hr
                                        DateTime temp = currentCellDateTime.AddHours(12);
                                        if (temp <= findNext)
                                        {
                                            worksheet.Cells[row, col].Value = temp;
                                            ApplyRedBorderStyle(row, col, ExcelBorderStyle.Thick, Color.Red);
                                        }
                                    }
                                }
                            }

                            if (col > 1) // middle and last cols only
                            {
                                object current = worksheet.Cells[row, col].Text;
                                DateTime.TryParse(current.ToString() ?? "", out DateTime currentCell);

                                object prevContent = worksheet.Cells[row, col - 1].Value;
                                if (prevContent != null)
                                {
                                    bool isPrevDouble = double.TryParse(prevContent.ToString(), out double prevContentAsDouble);
                                    if (isPrevDouble)
                                    {
                                        DateTime prevCol = DateTime.FromOADate(prevContentAsDouble);
                                        if (currentCell < prevCol)
                                        {
                                            ApplyCellFill(row, col, Color.Green);
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
        /// Apply a border to a cell with a color.
        /// </summary>
        /// <param name="row">The row the cell is on</param>
        /// <param name="col">The column the cell is on</param>
        /// <param name="borderStyle">The style of border to add around the edge</param>
        /// <param name="color">The color of the border around the edge</param>
        private void ApplyRedBorderStyle(int row, int col, ExcelBorderStyle borderStyle, Color color)
        {
            worksheet.Cells[row, col].Style.Border.Right.Style = borderStyle;
            worksheet.Cells[row, col].Style.Border.Left.Style = borderStyle;
            worksheet.Cells[row, col].Style.Border.Top.Style = borderStyle;
            worksheet.Cells[row, col].Style.Border.Bottom.Style = borderStyle;
            worksheet.Cells[row, col].Style.Border.Right.Color.SetColor(color);
            worksheet.Cells[row, col].Style.Border.Left.Color.SetColor(color);
            worksheet.Cells[row, col].Style.Border.Top.Color.SetColor(color);
            worksheet.Cells[row, col].Style.Border.Bottom.Color.SetColor(color);
        }

        /// <summary>
        /// For a specified cell given by the row and column.
        /// Change the background color with a solid fill.
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="color"></param>
        private void ApplyCellFill(int row, int col, Color color)
        {
            worksheet.Cells[row, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row, col].Style.Fill.BackgroundColor.SetColor(color);
        }
        #endregion
    }
}
