using System.Drawing;

namespace MedicorDataFormatter.Interfaces
{
    public interface IExcelStyler
    {
        /// <summary>
        /// Apply a background fill to a cell
        /// </summary>
        /// <param name="row">The row of the cell</param>
        /// <param name="col">The column of the cell</param>
        /// <param name="color">The color of the cell</param>
        void ApplyCellFill(int row, int col, Color color);
    }
}
