using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace MedicorDataFormatter.Excel
{
    /// <summary>
    /// Excel reader accesses an excel file and opens it 
    /// The class is used to retrieve the data
    /// </summary>
    public class ExcelReader
    {
        #region Properties
        private ExcelWorksheet worksheet;
        #endregion

        #region Constructors
        /// <summary>
        /// Find the excel file and retrieve a workbook by the given name
        /// </summary>
        /// <param name="path"></param>
        public ExcelReader(string path, string worksheetName)
        {
            FileInfo fileInfo = new FileInfo(path);

            ExcelPackage package = new ExcelPackage(fileInfo);

            // find the worksheet by name
            worksheet = package.Workbook.Worksheets
                .FirstOrDefault(x => x.Name.Equals(worksheetName, StringComparison.CurrentCultureIgnoreCase));
        }
        #endregion

        #region Read Excel Data
        public IList<object> ReadExcelWorksheet()
        {
            IList<object> excelData = new List<object>();

            int rows = worksheet.Dimension.Rows;
            int cols = worksheet.Dimension.Columns;

            for (int col = 1; col < cols; col++)
            {
                for (int row = 1; row < rows; row++)
                {
                    object content = worksheet.Cells[row, col].Value;

                    if (content == null) // the cell is blank, add correct string to it for the column
                    {
                        string insertValue = null;
                        switch (col)
                        {
                            case 1:
                                insertValue = "Time into theatre";
                                break;
                            case 2:
                                insertValue = "Time of Anaesthetic Start";
                                break;
                            case 3:
                                insertValue = "Time into Theatre";
                                break;
                            case 4:
                                insertValue = "Time out of Theatre";
                                break;
                            case 5:
                                insertValue = "Time into Recovery";
                                break;
                            case 6:
                                insertValue = "Time Out of Recovery";
                                break;
                            default:
                                insertValue = "Error";
                                break;
                        }

                        excelData.Add(insertValue);
                    }
                    else
                    {
                        string s = content.ToString();
                        bool isDouble = double.TryParse(s, out double numDateTime);

                        if (isDouble) // the value is a double
                        {
                            DateTime dateTime = DateTime.FromOADate(numDateTime);
                            excelData.Add(dateTime);
                        }
                        else // must be a header or someones entered a value not a datetime
                            continue;
                    }
                }
            }

            return excelData;
        }
        #endregion
    }
}
