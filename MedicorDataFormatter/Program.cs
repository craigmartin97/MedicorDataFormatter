using MedicorDataFormatter.Excel;
using OfficeOpenXml;
using System;
using System.IO;
using System.Linq;

namespace MedicorDataFormatter
{
    public class Program
    {
        public static void Main(string[] args)
        {
            ExcelReader excelReader = new ExcelReader(@"C:\Users\cmart\OneDrive\Desktop\MedicorDataFormatter\Dataset.xlsx", "Data");
            foreach(var v in excelReader.ReadExcelWorksheet())
            {
                Console.WriteLine(v);
            }
        }
    }
}
