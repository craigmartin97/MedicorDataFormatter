using MedicorDataFormatter.Excel;
using NUnit.Framework;
using OfficeOpenXml;
using System;
using System.IO;

namespace MedicorDataFormatter.Tests
{
    public class ExcelDataTests
    {

        /// <summary>
        /// Pass a invalid file path with a valid sheet
        /// Should throw exception
        /// </summary>
        [Test]
        public void SetupExcelFile_NullPath()
        {
            const string path = "";
            const string sheet = "Data";
            Assert.Throws<ArgumentNullException>(() => new ExcelData(path, sheet));
        }

        /// <summary>
        /// Pass a valid file with a invalid sheet
        /// Should throw exception
        /// </summary>
        [Test]
        public void SetupExcelFile_NullSheet()
        {
            const string path = "test";
            const string sheet = "";
            Assert.Throws<ArgumentNullException>(() => new ExcelData(path, sheet));
        }

        /// <summary>
        /// Supply a path that is invalid and cannot possibly be a valid file
        /// path. Should throw a FileNotFoundException
        /// </summary>
        [Test]
        public void SetupExcelFile_InvalidFileAddress()
        {
            const string path = "IAmAnInvalidAddrss";
            const string sheet = "Data";
            Assert.Throws<FileNotFoundException>(() => new ExcelData(path, sheet));
        }

        /// <summary>
        /// Supply a valid excel file but a worksheet that
        /// is extremely unlikely to be really.
        /// </summary>
        [Test]
        public void SetupExcelFile_InvalidWorksheet()
        {
            string path = System.IO.Path.GetTempPath() + Guid.NewGuid() + ".xlsx";
            string sheet = Guid.NewGuid().ToString();

            Assert.Throws<FileNotFoundException>(() => new ExcelData(path, sheet));
        }

        /// <summary>
        /// Create a temp excel file in the users temp folder
        /// Add a sheet with a name and then try and create it.
        /// Fianlly delete the file.
        /// </summary>
        [Test]
        public void SetupExcelFile_Valid()
        {
            ExcelPackage package = null;
            string path = Path.GetTempPath() + Guid.NewGuid() + ".xlsx";

            try
            {
                const string sheet = "Data";

                package = new ExcelPackage(new FileInfo(path));
                package.Workbook.Worksheets.Add(sheet);
                package.Save();

                ExcelData excelData = new ExcelData(path, sheet);

                Assert.IsNotNull(excelData);
                Assert.IsNotNull(excelData.Worksheet);
                Assert.IsNotNull(excelData.Package);
            }
            finally
            {
                if (package != null) // the package was created.
                {
                    File.Delete(path);
                }
            }
        }
    }
}