using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DNA.Common
{
    public class Sequence
    {
        public string Label { get; set; }
        public IList<Nucleotide> Data { get; set; }

        public Sequence()
        {
        }
    }
}
