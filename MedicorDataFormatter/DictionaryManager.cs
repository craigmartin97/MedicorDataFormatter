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
        /// Creates a dictionary with a key and value of ints
        /// </summary>
        /// <param name="section">Section to get from config file</param>
        /// <returns>Returns a dictionary of key int and value of int</returns>
        public Dictionary<int, int> GetIntDictionary(string section)
        {
            var columns = _configuration.GetSection(section).GetChildren();

            Dictionary<int, int> dictionary = new Dictionary<int, int>();
            foreach (var col in columns)
            {
                bool keyIsInt = int.TryParse(col.Key, out int key);
                bool valueIsInt = int.TryParse(col.Value, out int value);

                if (!keyIsInt || !valueIsInt) continue;

                dictionary.Add(key, value);
            }

            return dictionary;
        }
    }
}
