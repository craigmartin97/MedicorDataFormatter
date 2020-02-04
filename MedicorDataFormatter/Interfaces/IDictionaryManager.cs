using System.Collections.Generic;

namespace MedicorDataFormatter.Interfaces
{
    public interface IDictionaryManager
    {
        Dictionary<int, int> GetIntDictionary(string section);
    }
}