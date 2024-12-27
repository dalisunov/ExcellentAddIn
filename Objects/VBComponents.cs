using Microsoft.Vbe.Interop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcellentAddIn.Objects
{
    public class VBComponents
    {
        public VBComponents Components { get; private set; }

        public VBComponents(VBComponents components)
        {
            Components = components;
        }

        // Методы для управления VBA-модулями
    }
}
