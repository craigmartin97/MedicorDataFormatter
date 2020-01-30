using MedicorDataFormatter.Interfaces;
using Microsoft.Extensions.Configuration;
using System.Collections.Generic;
using System.Linq;

namespace MedicorDataFormatter
{
    /// <summary>
    /// DictionaryManager gets data from the configuration implementation
    /// and creates dictionaries based of it.
    /// </summary>
    public class DictionaryManager : IDictionaryManager
    {
        private readonly IConfiguration _configuration;

        public DictionaryManager(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        /// <summary>
        /// Returns a dictionary with a key string and string value.
        /// The method requires a section to be specified to be able to get from the config.
        /// </summary>
        /// <param name="section">The section in which to get the values from</param>
        /// <returns>Returns a dictionary of type string,stringn</returns>
        public Dictionary<string, string> GetDictionary(string section)
        {
            var columns = _configuration.GetSection(section).GetChildren();
            return columns.ToDictionary(col => col.Key, col => col.Value);
        }
    }
}
