using MedicorDataFormatter.Interfaces;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

namespace MedicorDataFormatter.Excel
{
    /// <summary>
    /// Excel styler class holds all logic to style the worksheet
    /// </summary>
    public class ExcelStyler : IExcelStyler
    {
        private readonly ExcelWorksheet _worksheet;

        /// <summary>
        /// ExcelStyler setup constructor.
        /// Takes a worksheet which is the sheet to engaged with to apply the styling too
        /// </summary>
        /// <param name="excelData"></param>
        public ExcelStyler(IExcelData excelData)
        {
            _worksheet = excelData.Worksheet;
        }

        /// <summary>
        /// Apply a border to a cell with a color.
        /// </summary>
        /// <param name="row">The row the cell is on</param>
        /// <param name="col">The column the cell is on</param>
        /// <param name="borderStyle">The style of border to add around the edge</param>
        /// <param name="color">The color of the border around the edge</param>
        public void ApplyBorderToCell(int row, int col, ExcelBorderStyle borderStyle, Color color)
        {
            _worksheet.Cells[row, col].Style.Border.Right.Style = borderStyle;
            _worksheet.Cells[row, col].Style.Border.Left.Style = borderStyle;
            _worksheet.Cells[row, col].Style.Border.Top.Style = borderStyle;
            _worksheet.Cells[row, col].Style.Border.Bottom.Style = borderStyle;
            _worksheet.Cells[row, col].Style.Border.Right.Color.SetColor(color);
            _worksheet.Cells[row, col].Style.Border.Left.Color.SetColor(color);
            _worksheet.Cells[row, col].Style.Border.Top.Color.SetColor(color);
            _worksheet.Cells[row, col].Style.Border.Bottom.Color.SetColor(color);
        }

        /// <summary>
        /// For a specified cell given by the row and column.
        /// Change the background color with a solid fill.
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="color"></param>
        public void ApplyCellFill(int row, int col, Color color)
        {
            _worksheet.Cells[row, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
            _worksheet.Cells[row, col].Style.Fill.BackgroundColor.SetColor(color);
        }

        /// <summary>
        /// Change the font color of the cell
        /// </summary>
        /// <param name="row">Row of the cell</param>
        /// <param name="col">Column of the cell</param>
        /// <param name="color">Color to make the font</param>
        public void ChangeFontColor(int row, int col, Color color)
            => _worksheet.Cells[row, col].Style.Font.Color.SetColor(color);
    }
}
