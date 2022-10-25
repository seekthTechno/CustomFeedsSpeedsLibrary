using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NXOpen;

namespace CustomFeedsSpeedsLibrary.Model
{
    public class LibraryNXOperation
    {
        public int operationType { get; set; }
        public double fZ { get; set; }
        public double Vc { get; set; }
        public double CutWidth { get; set; }
        public double CutStep { get; set; }  
        public Tag opTag { get; set; }
    }
}
