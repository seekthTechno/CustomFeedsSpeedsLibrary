using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NXOpen;
using NXOpen.CAM;

namespace CustomFeedsSpeedsLibrary.Model
{
    public class LibraryNXTool
    {
        public string toolName { get; set; }
        public string idInsert { get; set; }
        public string idTool { get; set; }
        public string holderName { get; set; }
        public double toolDiameter { get; set; }
        public int toolZCount { get; set; }
        public double toolOffset { get; set; }
        public double diamToOffset { get; set; }
        public Tool.Types ToolType { get; set; }       

        public NXObject.AttributeInformation[] toolAttributes { get; set; }               
    }
}
