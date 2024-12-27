using System.Collections;
using System.Collections.Generic;

namespace ExcellentAddIn.Core
{
    public class CollectionManager
    {
        public IList ConvertToIList<T>(IEnumerable<T> collection)
        {
            return new List<T>(collection);
        }
    }
}
