using System.Collections.Generic;

namespace ExcellentAddIn.Core
{
    public class DictionaryManager
    {
        public Dictionary<TKey, TValue> CreateDictionary<TKey, TValue>()
        {
            return new Dictionary<TKey, TValue>();
        }
    }
}
