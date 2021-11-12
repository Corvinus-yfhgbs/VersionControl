using _8.gyak.Abstractions;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _8.gyak.Entities
{
    public class PresentFactory : IToyFactory
    {
        public Color ribbon { get; set; }
        public Color box { get; set; }
        public Toy CreateNew()
        {
            return new Present(ribbon, box);
        }
    }
}
