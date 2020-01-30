using MedicorDataFormatter.Interfaces;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
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
        private readonly ExcelWorksheet _worksheet;

        /// <summary>
        /// The package, excel workbook, to work with
        /// </summary>
        private readonly ExcelPackage _package;

        /// <summary>
        /// Styler to apply styling to excel sheet
        /// </summary>
        private readonly IExcelStyler _styler;

        private readonly Dictionary<string, string> _nullCellDictionary;
        private readonly Dictionary<string, string> _columnDateTimeDictionary;
        #endregion

        #region Constructors

        /// <summary>
        /// Sets the excel styler and other injections needed
        /// </summary>
        /// <param name="excelData">The excel package and worksheet</param>
        /// <param name="excelStyler"></param>
        /// <param name="dictionaryManager"></param>
        public ExcelFormatter(IExcelData excelData, IExcelStyler excelStyler, IDictionaryManager dictionaryManager)
        {
            _worksheet = excelData.Worksheet;
            _package = excelData.Package;
            _styler = excelStyler;

            _nullCellDictionary = dictionaryManager.GetDictionary("Columns");
            _columnDateTimeDictionary = dictionaryManager.GetDictionary("BeforeColumns");
        }
        #endregion

        #region Read Excel Data
        public void FormatExcelHealthFile()
        {
            int rows = _worksheet.Dimension.Rows;
            int cols = _worksheet.Dimension.Columns;

            for (var col = 1; col <= cols; col++) //each column
            {
                for (var row = 1; row <= rows; row++) // each row
                {
                    // get the value from the cell
                    string content = GetTextFromCell(row, col);

                    if (string.IsNullOrWhiteSpace(content)) // the cell has nothing in, so add the phrase needed
                    {
                        InsertValueIntoNullCell(row, col);
                    }
                    else
                    {
                        // grab the text from the cell and try parse as DT
                        bool isDateTime = DateTime.TryParse(content, out DateTime currentCellDateTime);

                        if (isDateTime)
                        {
                            if (currentCellDateTime.Hour <= 9) // suspect 12 hour format carry on.
                            {
                                DateTime? findNext = null;
                                int nextCol = col; // nextCol, as in the prev or next col to the current col
                                bool isEnd;

                                if (cols == col) // last col
                                {
                                    do
                                    {
                                        isEnd = _worksheet.Dimension.Start.Column == nextCol;
                                        nextCol--;

                                        object prev = GetValueFromCell(row, nextCol);
                                        if (prev == null)
                                            continue;

                                        if (double.TryParse(prev.ToString(), out double nextDateAsNum))
                                            findNext = DateTime.FromOADate(nextDateAsNum);
                                    }
                                    while (findNext == null && !isEnd);

                                    if (findNext.HasValue)
                                    {
                                        if (findNext.Value.Hour <= 9) // if this is also less than 9, then it must be a morning (AM)
                                            continue;

                                        // add 12 hours to make 24 hr
                                        DateTime temp = currentCellDateTime.AddHours(12);
                                        if (findNext >= temp)
                                        {
                                            _worksheet.Cells[row, col].Value = temp;
                                            _styler.ApplyBorderToCell(row, col, ExcelBorderStyle.Thick, Color.Red);
                                        }
                                    }
                                }
                                else // first or middle cols
                                {
                                    do
                                    {
                                        isEnd = cols == nextCol; // true
                                        nextCol++;

                                        object next = GetValueFromCell(row, nextCol); ;
                                        if (next == null)
                                            continue;

                                        if (double.TryParse(next.ToString(), out double nextDateAsNum))
                                            findNext = DateTime.FromOADate(nextDateAsNum);
                                    }
                                    while (findNext == null && !isEnd);

                                    if (findNext.HasValue)
                                    {
                                        if (findNext.Value.Hour <= 9) // if this is also less than or equal to 9, then it must be a morning (AM)
                                            continue;

                                        // add 12 hours to make 24 hr
                                        DateTime temp = currentCellDateTime.AddHours(12);
                                        if (temp <= findNext)
                                        {
                                            _worksheet.Cells[row, col].Value = temp;
                                            _styler.ApplyBorderToCell(row, col, ExcelBorderStyle.Thick, Color.Red);

                                            /*
                                             * start at the current column and loop backwards columns ensure, no others are left
                                             * as 12 hr formats
                                             */
                                            if (col > _worksheet.Dimension.Start.Column)
                                            {
                                                for (int i = col - 1; i >= _worksheet.Dimension.Start.Column; i--)
                                                {
                                                    if (DateTime.TryParse(_worksheet.Cells[row, i].Text, out DateTime prevDateTime))
                                                    {
                                                        DateTime tempPrevDateTime = prevDateTime.AddHours(12);
                                                        if (tempPrevDateTime <= temp) // added tweleve hours on and its still less than temp, so must now be 24hr format
                                                        {
                                                            _worksheet.Cells[row, i].Value = tempPrevDateTime;
                                                            _styler.ApplyBorderToCell(row, i, ExcelBorderStyle.Thick, Color.Blue);
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }


                            CheckIfDateTimeIsBefore(row, col);
                        }
                    }
                }
            }

            _package.Save(); // save the excel file. Could throw InvalidOperationException if open in other program!!!
        }
        #endregion

        #region Null String Value

        /// <summary>
        /// Gets the null reference phrase based upon the column index
        /// and inserts the value into the current null cell.
        ///
        /// Use the current columns first row as the header column.
        /// Get the text and retrieve and get the value from the dictionary with the same key
        /// </summary>
        /// <param name="row">The row of the null cell</param>
        /// <param name="col">The column of the null cell</param>
        private void InsertValueIntoNullCell(int row, int col)
        {
            // get the current cols header. Use the first row and the current column index.
            string colHeader = _worksheet.Cells[_worksheet.Dimension.Start.Row, col].Text;
            if (!string.IsNullOrWhiteSpace(colHeader)) // actually got a header that is a string
            {
                if (GetValueFromDictionary(_nullCellDictionary, colHeader, out string inputValue))
                {
                    // got header value can insert into null cell.
                    _worksheet.Cells[row, col].Value = inputValue;
                    _styler.ApplyBorderToCell(row, col, ExcelBorderStyle.Thick, Color.DeepSkyBlue);
                }
            }
        }

        #endregion

        #region DateTime Before
        /// <summary>
        /// Check if the date and time of the cell is before another one.
        /// If it is then highlight the cell.
        /// </summary>
        private void CheckIfDateTimeIsBefore(int row, int col)
        {
            int firstRow = _worksheet.Dimension.Start.Row;

            // get the value from the current cell
            if (GetDateTimeFromString(GetTextFromCell(row, col), out DateTime currentCell))
            {
                /*
                 * get the current cells header. First row, current col
                 * use the current headers value as key, and search inn the dictionary to get the comparision columns header title
                 */
                if (GetValueFromDictionary(_columnDateTimeDictionary, GetTextFromCell(firstRow, col), out string compareColHeader))
                {
                    int colIndexOfComparison = 0;
                    // start at the first column
                    for (int i = 1; i <= _worksheet.Dimension.Columns; i++)
                    {
                        // compare the cells text to the key's value.
                        if (GetTextFromCell(firstRow, i).Equals(compareColHeader)) // found the cell that matches
                        {
                            colIndexOfComparison = i;
                            break;
                        }
                    }

                    if (colIndexOfComparison > 0) // got index value of the comparison column.
                    {
                        if (GetDateTimeFromString(GetTextFromCell(row, colIndexOfComparison), out DateTime compareCellDateTime))
                        {
                            if (currentCell < compareCellDateTime)
                            {
                                // the current cell is less than the compare cell's value. Apply formatting
                                _styler.ApplyCellFill(row, col, Color.Green);
                                _styler.ChangeFontColor(row, col, Color.White);
                            }
                        }
                    }
                }
            }

            /*
             * This is old code, 
             * leave here for now. Even though in version control
             */

            //if (col > 1) // middle and last cols only
            //{

            //    bool currentDateTime = DateTime.TryParse(GetTextFromCell(row, col), out DateTime currentCell);
            //    bool prevDateTime = DateTime.TryParse(GetTextFromCell(row, col - 1), out DateTime prevContent);

            //    if (currentDateTime && prevDateTime)
            //    {
            //        if (currentCell < prevContent)
            //        {
            //            _styler.ApplyCellFill(row, col, Color.Green);
            //            _styler.ChangeFontColor(row, col, Color.White);
            //        }
            //    }
            //}
        }
        #endregion

        #region Helpers
        /// <summary>
        /// Get the value from the workbook cell.
        /// </summary>
        /// <param name="row">The row of the cell</param>
        /// <param name="col">The column of the cell</param>
        /// <returns>Returns an object from the cell</returns>
        private object GetValueFromCell(int row, int col) => _worksheet.Cells[row, col].Value;

        /// <summary>
        /// Gets the text from the workbook cell
        /// </summary>
        /// <param name="row">The row of the cell</param>
        /// <param name="col">The column of the cell</param>
        /// <returns>Returns the cells value as a string</returns>
        private string GetTextFromCell(int row, int col) => _worksheet.Cells[row, col].Text;

        /// <summary>
        /// Get a value from a dictionary
        /// </summary>
        /// <param name="dictionary">Dictionary to get value from</param>
        /// <param name="key">key value to search for</param>
        /// <param name="retrievedVal">The value to be retrieved</param>
        /// <returns></returns>
        private bool GetValueFromDictionary(Dictionary<string, string> dictionary, string key, out string retrievedVal)
            => dictionary.TryGetValue(key, out retrievedVal);

        /// <summary>
        /// Try and parse a string and output a datetime
        /// </summary>
        /// <param name="value">Value to try and parse</param>
        /// <param name="dateTime">date to output in response</param>
        /// <returns>Returns true if the operation was successful</returns>
        private bool GetDateTimeFromString(string value, out DateTime dateTime)
            => DateTime.TryParse(value, out dateTime);

        #endregion
    }
}
