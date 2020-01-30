using MedicorDataFormatter.Excel;
using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;

namespace MedicorDataFormatter
{
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                ExcelFormatter excelReader = new ExcelFormatter(@"E:\Medicor\MedicorDataFormatter\Dataset.xlsx", "Data");
                excelReader.FormatExcelHealthFile();
            }
            catch (FileNotFoundException ex)
            {
                Debug.WriteLine("The file or worksheet could not be found!!");
                Console.WriteLine(ex.Message);
            }
            catch (ArgumentNullException ex)
            {
                Debug.WriteLine("The file path or workbook name are invalid, possibly null or blank");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
