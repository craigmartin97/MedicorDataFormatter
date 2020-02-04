using MedicorDataFormatter.Interfaces;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;

namespace MedicorDataFormatter.Excel
{
    /// <summary>
    /// Excel reader accesses an excel file and opens it 
    /// The class is used to retrieve the data
    /// </summary>
    public class ExcelFormatter : IExcelFormatter
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

        /// <summary>
        /// A dictionary to match the column header to a key
        /// to output a null value to insert, when the cell is null.
        /// </summary>
        private readonly Dictionary<int, int> _nullCellDictionary;

        /// <summary>
        /// A dictionary to hold keys and values.
        /// Used to get the columns header and match with a key in the dictionary.
        /// Outputs the other column to compare with
        /// </summary>
        private readonly Dictionary<int, int> _columnDateTimeDictionary;
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

            _nullCellDictionary = dictionaryManager.GetIntDictionary("Columns");

            if (_nullCellDictionary == null)
                throw new NullReferenceException("The dictionary for columns is null. Ensure the configuration file is correct");

            _columnDateTimeDictionary = dictionaryManager.GetIntDictionary("BeforeColumns");

            if (_columnDateTimeDictionary == null)
                throw new NullReferenceException("The dictionary for before check columns is null. Ensure the configuration file is correct");
        }
        #endregion

        #region Read Excel Data
        /// <summary>
        /// Format the excel file.
        /// Loops through each cell and gets the value from it.
        /// If the cell is null then a phrase is inserted
        /// If the cell has content the date is checked if it is 24hr format and a valid time based on other cols
        /// </summary>
        public void FormatExcelHealthFile()
        {
            // each row, start at the second row to skip the row headers.
            for (int row = _worksheet.Dimension.Start.Row + 1; row <= _worksheet.Dimension.Rows; row++)
            {
                for (int col = _worksheet.Dimension.Start.Column; col <= _worksheet.Dimension.Columns; col++) //each column
                {
                    InsertValueIntoNullCell(row, col); // null value
                    ChangeTimeFormat(row, col); // 12hr to 24hr
                    CheckIfDateTimeIsBefore(row, col); // cell highlighting
                }
            }

            _package.Save(); // save the excel file. Could throw InvalidOperationException if excel sheet open in other program.
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
            string currentCell = GetTextFromCell(row, col);
            if (!string.IsNullOrWhiteSpace(currentCell)) return;

            bool receivedCompareColIndex = GetColIndexFromDictionary(_nullCellDictionary, col, out int compareColIndex);
            if (!receivedCompareColIndex) return;

            bool isDateTime = GetDateTimeFromString(GetTextFromCell(row, compareColIndex), out DateTime compareDateTime);
            if (!isDateTime) return;

            // got header value can insert into null cell.
            InsertValueIntoCell(row, col, compareDateTime);
            _styler.ApplyBorderToCell(row, col, ExcelBorderStyle.Thick, Color.DeepSkyBlue);
        }

        #endregion

        #region DateTime Before
        /// <summary>
        /// Check if the date and time of the cell is before another one.
        /// If it is then highlight the cell.
        ///
        /// NOTE: Spec said for col "Surgery finish time" to check with
        /// column "Time into theatre", unsure if this is correct as it doesnt follow the pattern
        /// of the other columns. Should it be checking with "Surgery start time"? The col to the left.
        /// Easy changeable, just go to appsettings.json and edit the corresponding value.
        /// </summary>
        private void CheckIfDateTimeIsBefore(int row, int col)
        {
            // get current cells value
            bool isCurrentDateTime = GetDateTimeFromString(GetTextFromCell(row, col), out DateTime currentCell);
            if (!isCurrentDateTime) return;

            // get the current columns comparisons index
            bool isCompareColInt = GetColIndexFromDictionary(_columnDateTimeDictionary, col, out int compareColIndex);
            if (!isCompareColInt) return;

            // get the compare columns cell for the relevant row
            bool isCompareColDateTime =
                GetDateTimeFromString(GetTextFromCell(row, compareColIndex), out DateTime compareColDateTime);
            if (!isCompareColDateTime) return;

            if (currentCell < compareColDateTime)
            {
                // the current cell is less than the compare cell's value. Apply formatting
                _styler.ApplyCellFill(row, col, Color.Green);
                _styler.ChangeFontColor(row, col, Color.White);
            }
        }
        #endregion

        #region Change clock formatting
        /// <summary>
        /// Change the time from 12 hr format to 24hr format if appropriate
        /// </summary>
        /// <param name="row">The row of the cell</param>
        /// <param name="col">The column of the cell</param>
        private void ChangeTimeFormat(int row, int col)
        {
            bool currentIsDate = GetDateTimeFromString(GetTextFromCell(row, col), out DateTime currentCellDateTime);
            if (currentIsDate && currentCellDateTime.Hour <= 12) // 12 or less means it could be 12 to 24 hr conversion needed
            {
                if (col == _worksheet.Dimension.End.Column) // in last col, there is no next col use prev
                {
                    // get datetime from prev cell
                    if (!GetDateTimeFromString(GetTextFromCell(row, col - 1),
                        out DateTime prevCellDateTime)) return;

                    currentCellDateTime = currentCellDateTime.AddHours(12);

                    if (prevCellDateTime < currentCellDateTime) return;

                    InsertValueIntoCell(row, col, currentCellDateTime);
                    _styler.ApplyBorderToCell(row, col, ExcelBorderStyle.Thick, Color.Red);
                }
                else if (col == _worksheet.Dimension.Start.Column) // first column
                {
                    // get datetime from next cell
                    if (!GetDateTimeFromString(GetTextFromCell(row, col + 1),
                        out DateTime nextCellDateTime)) return;

                    // add twelve hours on to current
                    currentCellDateTime = currentCellDateTime.AddHours(12);

                    if (currentCellDateTime > nextCellDateTime) return; // can't possibly be correct. Don't edit on the sheet

                    InsertValueIntoCell(row, col, currentCellDateTime);
                    _styler.ApplyBorderToCell(row, col, ExcelBorderStyle.Thick, Color.Red);
                }
                else // middle cells
                {
                    // get datetime from prev cell
                    if (!GetDateTimeFromString(GetTextFromCell(row, col - 1),
                        out DateTime prevCellDateTime)) return;

                    // get datetime from next cell
                    if (!GetDateTimeFromString(GetTextFromCell(row, col + 1),
                        out DateTime nextCellDateTime)) return;

                    /*
                     * Current cell is less than the previous one.
                     * That will mean it could be a 12 hr time that needs converting.
                     */
                    if (currentCellDateTime < prevCellDateTime)
                    {
                        // add twelve hours on to current
                        currentCellDateTime = currentCellDateTime.AddHours(12);

                        // the date and time is now between the two dates it must be converted to 24hr clock
                        if (currentCellDateTime >= prevCellDateTime && currentCellDateTime <= nextCellDateTime)
                        {
                            InsertValueIntoCell(row, col, currentCellDateTime);
                            _styler.ApplyBorderToCell(row, col, ExcelBorderStyle.Thick, Color.Red);
                        }
                    }
                }
            }
        }
        #endregion 

        #region Helpers
        /// <summary>
        /// Gets the text from the workbook cell
        /// </summary>
        /// <param name="row">The row of the cell</param>
        /// <param name="col">The column of the cell</param>
        /// <returns>Returns the cells value as a string</returns>
        private string GetTextFromCell(int row, int col)
            => _worksheet.Cells[row, col].Text;

        /// <summary>
        /// Get a value from a dictionary
        /// </summary>
        /// <param name="dictionary">Dictionary to get value from</param>
        /// <param name="key">key value to search for</param>
        /// <param name="colValue">The column value</param>
        /// <returns>Returns an integer returned from the dictionary</returns>
        private bool GetColIndexFromDictionary(Dictionary<int, int> dictionary, int key, out int colValue)
            => dictionary.TryGetValue(key, out colValue);

        /// <summary>
        /// Try and parse a string and output a datetime
        /// </summary>
        /// <param name="value">Value to try and parse</param>
        /// <param name="dateTime">date to output in response</param>
        /// <returns>Returns true if the operation was successful</returns>
        private bool GetDateTimeFromString(string value, out DateTime dateTime)
            => DateTime.TryParse(value, out dateTime);

        /// <summary>
        /// Insert a value into the cell
        /// </summary>
        /// <param name="row">Row of the cell</param>
        /// <param name="col">Column of the cell</param>
        /// <param name="value">Value to insert</param>
        private void InsertValueIntoCell(int row, int col, object value)
            => _worksheet.Cells[row, col].Value = value;

        #endregion
    }
}
