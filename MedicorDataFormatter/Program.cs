using MedicorDataFormatter.Excel;
using System;
using System.Diagnostics;
using System.IO;

namespace MedicorDataFormatter
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var watch = System.Diagnostics.Stopwatch.StartNew();

            // get root argument, for device
            string root = null;
            for (int i = 0; i < args.Length; i++)
            {
                if (args[i].Equals("-root"))
                {
                    root = args[i + 1];
                }
            }

            // validate root found
            if (string.IsNullOrWhiteSpace(root))
            {
                Console.WriteLine("You must supply a root command line argument!");
                return;
            }

            // try and format the excel sheet
            try
            {
                string path = string.Format(@"{0}Dataset.xlsx", root);
                ExcelFormatter excelReader = new ExcelFormatter(path, "Data");
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
            catch(InvalidOperationException ex)
            {
                Debug.WriteLine("Unable to save the workbook. Is it open in another program?");
                Console.WriteLine(ex.Message);
                Console.WriteLine("Unable to save the workbook. If open in other program, close it down");
            }

            watch.Stop();
            Console.WriteLine("Execution Time: " + watch.ElapsedMilliseconds + "ms");

        }
    }
}
