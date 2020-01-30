using System.Collections.Generic;

namespace MedicorDataFormatter.Interfaces
{
    public interface IDictionaryManager
    {
        Dictionary<string, string> GetDictionary(string section);
    }
}