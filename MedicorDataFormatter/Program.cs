using MedicorDataFormatter.Excel;
using MedicorDataFormatter.Interfaces;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Diagnostics;
using System.IO;

namespace MedicorDataFormatter
{
    public class Program
    {
        /// <summary>
        /// The services provider holding the items for dp injection
        /// </summary>
        private static IServiceProvider _serviceProvider;

        private static IConfiguration _configuration;

        public static void Main(string[] args)
        {
            // start timer watch for performance analysis
            var watch = Stopwatch.StartNew();

            SetupConfig();

            FormatExcelFile();

            watch.Stop();
            Console.WriteLine("Execution Time: " + watch.ElapsedMilliseconds + "ms");
        }

        #region Excel File
        private static void FormatExcelFile()
        {
            // try and format the excel sheet
            try
            {
                IExcelFormatter excelReader = _serviceProvider.GetService<IExcelFormatter>();
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
            catch (InvalidOperationException ex)
            {
                Debug.WriteLine("Unable to save the workbook. Is it open in another program?");
                Console.WriteLine(ex.Message);
                Console.WriteLine("Unable to save the workbook. If open in other program, close it down");
            }
        }
        #endregion

        #region Config
        /// <summary>
        /// Setup the configuration settings.
        /// Add in the config file and run the dependency injection code.
        /// </summary>
        private static void SetupConfig()
        {
            /*
             * add in the app settings file
             * The columns headers on the dataset sheet.
             * The key is the header text at the top. The value is the phrase
             * to insert upon null / blankness in the cell
             */
            _configuration = new ConfigurationBuilder()
                .SetBasePath(Environment.CurrentDirectory)
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: false)
                .Build(); // build the config and store for usage

            // configure the services to dependency inject
            ServiceCollection serviceCollection = new ServiceCollection();
            ConfigureServices(serviceCollection);
            _serviceProvider = serviceCollection.BuildServiceProvider();
        }

        /// <summary>
        /// Setup the configuration service for dependency injection
        /// </summary>
        /// <param name="serviceCollection">The collection storing the items for dp injection</param>
        private static void ConfigureServices(IServiceCollection serviceCollection)
        {
            serviceCollection.AddSingleton(_configuration); // add config 

            serviceCollection.AddSingleton<IDictionaryManager, DictionaryManager>(); // add dictionary manager to get items from dict

            serviceCollection.AddSingleton<IExcelData, ExcelData>(x
                => new ExcelData($@"{_configuration["FileRoot"]}{_configuration["FileName"]}.xlsx",
                    _configuration["WorksheetName"]));

            serviceCollection.AddSingleton<IExcelStyler, ExcelStyler>();

            serviceCollection.AddSingleton<IExcelFormatter, ExcelFormatter>();
        }
        #endregion  
    }
}
