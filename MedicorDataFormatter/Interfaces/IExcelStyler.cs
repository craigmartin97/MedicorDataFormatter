using OfficeOpenXml.Style;
using System.Drawing;

namespace MedicorDataFormatter.Interfaces
{
    /// <summary>
    /// Interface for styling excel spreadsheet cells
    /// </summary>
    public interface IExcelStyler
    {
        /// <summary>
        /// Apply a border to a cell on the worksheet
        /// </summary>
        /// <param name="row">Row of cell</param>
        /// <param name="col">Column of cell</param>
        /// <param name="borderStyle">Border style to apply</param>
        /// <param name="color">Color of the border</param>
        void ApplyBorderToCell(int row, int col, ExcelBorderStyle borderStyle, Color color);

        /// <summary>
        /// For a specified cell given by the row and column.
        /// Change the background color with a solid fill.
        /// </summary>
        /// <param name="row">The row of the cell</param>
        /// <param name="col">The column of the cell</param>
        /// <param name="color">The color to fill the cell with</param>
        void ApplyCellFill(int row, int col, Color color);

        /// <summary>
        /// Change the color of the font for the cell.
        /// </summary>
        /// <param name="row">Row of the cell</param>
        /// <param name="col">Column of the cell</param>
        /// <param name="color">Color of the font to change to</param>
        void ChangeFontColor(int row, int col, Color color);
    }
}
