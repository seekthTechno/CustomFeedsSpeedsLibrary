using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NXOpen;
using NXOpen.CAM;
using CsvHelper.Configuration.Attributes;

namespace CustomFeedsSpeedsLibrary.Model
{
    public class LibraryElement
    {
        [Index(0)]
        public string ToolType { get; set; }

        [Index(1)]
        public string ToolSubType { get; set; }

        [Index(2)]
        public string toolName { get; set; }

        [Index(3)]
        public string idInsert { get; set; }

        [Index(4)]
        public string idTool { get; set; }

        [Index(5)]
        public string holderName { get; set; }

        [Index(6)]
        public string toolDiameter { get; set; }

        [Index(7)]
        public string toolZCount { get; set; }

        [Index(8)]
        public string toolOffset { get; set; }

        [Index(9)]
        public string diamToOffset { get; set; }
               
        [Index(10)]
        public string fZ { get; set; }

        [Index(11)]
        public string Vc { get; set; }

        [Index(12)]
        public string CutWidth { get; set; }

        [Index(13)]
        public string CutStep { get; set; }

        [Index(14)]
        public string operationType { get; set; }

        [Index(15)]
        public string globalMachineName { get; set; }

        [Index(16)]
        public string globalPartMat { get; set; }

        [Index(17)]
        public string globalPartHard { get; set; }

        [Index(18)]
        public string globalPartFastening { get; set; }

        [Index(19)]
        public string userName { get; set; }

        [Index(20)]
        public string partName { get; set; }

        [Index(21)]
        public string opGroupName { get; set; }

        [Index(22)]
        public string DateTime { get; set; }

    }
}
