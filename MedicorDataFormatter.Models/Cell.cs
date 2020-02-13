namespace MedicorDataFormatter.Models
{
    public class Cell<T>
    {
        public int Row { get; set; }
        public int Column { get; set; }
        public T Value { get; set; }

        #region Constructors
        public Cell() { }

        /// <summary>
        /// Create a new cell object with given data
        /// </summary>
        /// <param name="row">Row of the cell</param>
        /// <param name="col">Column of the cell</param>
        /// <param name="obj">Type object</param>
        public Cell(int row, int col, T obj)
        {
            Row = row;
            Column = col;
            Value = obj;
        }
        #endregion
    }
}