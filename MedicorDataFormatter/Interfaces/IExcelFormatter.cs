using System;
using System.Collections.Generic;
using MedicorDataFormatter.Models;

namespace MedicorDataFormatter.Interfaces
{
    public interface IExcelFormatter
    {
        IList<Cell<DateTime?>> Changes { get; }
        void FormatExcelHealthFile();
    }
}