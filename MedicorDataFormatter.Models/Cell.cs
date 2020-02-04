using System;

namespace MedicorDataFormatter.Models
{
    public class Cell<T> 
    {
        public int Row { get; set; }
        public int Column { get; set; }
        public T Value { get; set; }
    }
}