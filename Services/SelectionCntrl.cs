using System;
using System.Collections.Generic;
using System.Linq;
using NXOpen;
using NXOpen.CAM;
using CustomFeedsSpeedsLibrary.Model;
using NXOpen.Utilities;
using NXOpen.UF;
using CustomFeedsSpeedsLibrary.Services;

namespace CustomFeedsSpeedsLibrary
{
    public class SelectionCntrl
    {
        private LibraryRecordUpload lrUpload;        
        
        private static Session theSession = Session.GetSession();
        private static UFSession ufSession = UFSession.GetUFSession();
        private static Part workPart = theSession.Parts.Work;        
        
        public List<CAMObject> operationList { get; set; } = new List<CAMObject>();
        public List<LibraryElement> ElementList { get; set; } = new List<LibraryElement>();

        private static string[] opTypeFilter = new string[] { "500", "510", "520", "530", "540", "560" };

        public void getSelectedandSave(List<TaggedObject> objects, DataService _ds)
        {
            Tag[] tList = null;
            int count;
                     
            if (objects == null) return;

            foreach (var obj in objects)
            {
                if (obj is NCGroup)
                {                    
                    ufSession.Ncgroup.AskMemberList(obj.Tag, out count, out tList);
                    foreach (Tag t in tList)
                    {
                        operationList.Add(NXObjectManager.Get(t) as CAMObject);
                    }                        
                }

                if (obj is Operation)
                {
                    operationList.Add(NXObjectManager.Get(obj.Tag) as CAMObject);                    
                }
            }                          

            foreach (Operation op in operationList.ToList())
            {
                LibraryElement libElem = new LibraryElement();
                libElem.globalMachineName = _ds.machineName;
                libElem.globalPartMat = _ds.partMaterial;
                libElem.globalPartHard = _ds.partHardness.ToString();
                libElem.globalPartFastening = _ds.partFastening.ToString();
                
                GetOpParameters(op, libElem, workPart);

                if (libElem.operationType == null || opTypeFilter.Contains(libElem.operationType))
                {
                    //ElementList.Remove(le);
                    operationList.Remove(op);
                    continue;
                }

                GetToolParameters(op, libElem, workPart);

                libElem.DateTime = DateTime.Today.ToShortDateString();
                libElem.userName = System.Environment.UserName;

                if (!String.IsNullOrEmpty(workPart.GetUserAttributeAsString("НОМЕР_ДЕТАЛИ", NXObject.AttributeType.String, -1))) libElem.partName = workPart.GetUserAttributeAsString("НОМЕР_ДЕТАЛИ", NXObject.AttributeType.String, -1);
                else libElem.partName = "";

                libElem.opGroupName = op.Name;

                ElementList.Add(libElem);
            }
            // Flush operations List
            operationList.Clear();

            lrUpload = new LibraryRecordUpload();
            lrUpload.save2file(ElementList);

            // Flush Element List when done
            ElementList.Clear();
        }
        
        /*  For later release
        public void getSelectedandLoad(List<TaggedObject> objects)
        {
            Tag[] tList = null;
            int count;

            if (objects == null) return;

            foreach (var obj in objects)
            {
                if (obj is NCGroup)
                {
                    ufSession.Ncgroup.AskMemberList(obj.Tag, out count, out tList);
                    foreach (Tag t in tList)
                    {
                        operationList.Add(NXObjectManager.Get(t) as CAMObject);
                    }
                }

                if (obj is Operation)
                {
                    operationList.Add(NXObjectManager.Get(obj.Tag) as CAMObject);
                }
            }

            Tag toolTag;

            foreach (Operation op in operationList)
            {
                ufSession.Oper.AskCutterGroup(op.Tag, out toolTag);
                LibraryRecord lRec = new LibraryRecord();
                GetSmthng(op, lRec, workPart);
            }
        }
       
        public int cycleOpList(List<CAMObject> opList)
        {
            foreach (NXObject op in opList)
            {

            }
            return 42;
        }
        */

        public void GetToolParameters(Operation op, LibraryElement _le, Part workPart)
        {
            Tag toolTag;
            Tool.Types toolType;
            Tool.Subtypes toolSubType;
            ufSession.Oper.AskCutterGroup(op.Tag, out toolTag);
            
            var _tool = NXObjectManager.Get(toolTag) as Tool;

            _tool.GetTypeAndSubtype(out toolType, out toolSubType);
                        
            _le.ToolType = toolType.ToString();
            _le.ToolSubType = toolSubType.ToString();
            _le.toolName = string.Format(_tool.Name);
            NXObject.AttributeInformation[] toolAttributes = _tool.GetUserAttributes();
            
            if (toolAttributes.Any(x => x.Title == "ID_INSERT"))
            {
                var ins = toolAttributes
                     .Where(x => x.Title == "ID_INSERT")
                     .First();

                _le.idInsert = ins.StringValue;
            }
            else { _le.idInsert = ""; }

            if (toolAttributes.Any(x => x.Title == "ID_TOOL"))
            {
                var tl = toolAttributes
                  .Where(x => x.Title == "ID_TOOL")
                  .First();

                _le.idTool = tl.StringValue;
            }
            else { _le.idTool = ""; }
            
            MillingToolBuilder millingToolBuilder;
            switch (toolType)
            {
                case Tool.Types.Mill:
                    millingToolBuilder = workPart.CAMSetup.CAMGroupCollection.CreateMillToolBuilder(_tool);
                    _le.holderName = millingToolBuilder.TlHolderDescription;
                    _le.toolDiameter = string.Format("{0:0.0}", millingToolBuilder.TlDiameterBuilder.Value);
                    _le.toolZCount = millingToolBuilder.TlNumFlutesBuilder.Value.ToString();
                    _le.toolOffset = string.Format("{0:0.0}", (millingToolBuilder.TlHeightBuilder.Value + millingToolBuilder.TaperedShankLengthBuilder.Value) - millingToolBuilder.TlZMountBuilder.Value);
                    _le.diamToOffset = string.Format("{0:0.0}", Math.Round((millingToolBuilder.TlHeightBuilder.Value +
                        millingToolBuilder.TaperedShankLengthBuilder.Value) -
                        millingToolBuilder.TlZMountBuilder.Value)
                        / millingToolBuilder.TlDiameterBuilder.Value, 2);
                    break;

                case Tool.Types.Barrel:
                    millingToolBuilder = workPart.CAMSetup.CAMGroupCollection.CreateBarrelToolBuilder(_tool);
                    _le.holderName = millingToolBuilder.TlHolderDescription;
                    _le.toolDiameter = string.Format("{0:0.0}", millingToolBuilder.TlDiameterBuilder.Value);
                    _le.toolZCount = millingToolBuilder.TlNumFlutesBuilder.Value.ToString();
                    _le.toolOffset = string.Format("{0:0.0}", (millingToolBuilder.TlHeightBuilder.Value + millingToolBuilder.TaperedShankLengthBuilder.Value) - millingToolBuilder.TlZMountBuilder.Value);
                    _le.diamToOffset = string.Format("{0:0.0}", Math.Round((millingToolBuilder.TlHeightBuilder.Value +
                        millingToolBuilder.TaperedShankLengthBuilder.Value) -
                        millingToolBuilder.TlZMountBuilder.Value)
                        / millingToolBuilder.TlDiameterBuilder.Value, 2);
                    break;

                case Tool.Types.Tcutter:
                    millingToolBuilder = workPart.CAMSetup.CAMGroupCollection.CreateTToolBuilder(_tool);
                    _le.holderName = millingToolBuilder.TlHolderDescription;
                    _le.toolDiameter = string.Format("{0:0.0}", millingToolBuilder.TlDiameterBuilder.Value);
                    _le.toolZCount = millingToolBuilder.TlNumFlutesBuilder.Value.ToString();
                    _le.toolOffset = string.Format("{0:0.0}", (millingToolBuilder.TlHeightBuilder.Value + millingToolBuilder.TaperedShankLengthBuilder.Value) - millingToolBuilder.TlZMountBuilder.Value);
                    _le.diamToOffset = string.Format("{0:0.0}", Math.Round((millingToolBuilder.TlHeightBuilder.Value +
                        millingToolBuilder.TaperedShankLengthBuilder.Value) -
                        millingToolBuilder.TlZMountBuilder.Value)
                        / millingToolBuilder.TlDiameterBuilder.Value, 2);
                    break;

                case Tool.Types.Drill:
                switch (toolSubType) { 
                    case Tool.Subtypes.DrillStandard:
                            millingToolBuilder = workPart.CAMSetup.CAMGroupCollection.CreateDrillStdToolBuilder(_tool);
                            _le.holderName = millingToolBuilder.TlHolderDescription;
                            _le.toolDiameter = string.Format("{0:0.0}", millingToolBuilder.TlDiameterBuilder.Value);
                            _le.toolZCount = millingToolBuilder.TlNumFlutesBuilder.Value.ToString();
                            _le.toolOffset = string.Format("{0:0.0}", (millingToolBuilder.TlHeightBuilder.Value + millingToolBuilder.TaperedShankLengthBuilder.Value) - millingToolBuilder.TlZMountBuilder.Value);
                            _le.diamToOffset = string.Format("{0:0.0}", Math.Round((millingToolBuilder.TlHeightBuilder.Value +
                                millingToolBuilder.TaperedShankLengthBuilder.Value) -
                                millingToolBuilder.TlZMountBuilder.Value)
                                / millingToolBuilder.TlDiameterBuilder.Value, 2);
                            break;

                    case Tool.Subtypes.DrillCenterBell:
                            millingToolBuilder = workPart.CAMSetup.CAMGroupCollection.CreateDrillCenterBellToolBuilder(_tool);
                            _le.holderName = millingToolBuilder.TlHolderDescription;
                            _le.toolDiameter = string.Format("{0:0.0}", millingToolBuilder.TlDiameterBuilder.Value);
                            _le.toolZCount = millingToolBuilder.TlNumFlutesBuilder.Value.ToString();
                            _le.toolOffset = string.Format("{0:0.0}", (millingToolBuilder.TlHeightBuilder.Value + millingToolBuilder.TaperedShankLengthBuilder.Value) - millingToolBuilder.TlZMountBuilder.Value);
                            _le.diamToOffset = string.Format("{0:0.0}", Math.Round((millingToolBuilder.TlHeightBuilder.Value +
                                millingToolBuilder.TaperedShankLengthBuilder.Value) -
                                millingToolBuilder.TlZMountBuilder.Value)
                                / millingToolBuilder.TlDiameterBuilder.Value, 2);
                            break;

                    case Tool.Subtypes.DrillCountersink:
                            millingToolBuilder = workPart.CAMSetup.CAMGroupCollection.CreateDrillCtskToolBuilder(_tool);
                            _le.holderName = millingToolBuilder.TlHolderDescription;
                            _le.toolDiameter = string.Format("{0:0.0}", millingToolBuilder.TlDiameterBuilder.Value);
                            _le.toolZCount = millingToolBuilder.TlNumFlutesBuilder.Value.ToString();
                            _le.toolOffset = string.Format("{0:0.0}", (millingToolBuilder.TlHeightBuilder.Value + millingToolBuilder.TaperedShankLengthBuilder.Value) - millingToolBuilder.TlZMountBuilder.Value);
                            _le.diamToOffset = string.Format("{0:0.0}", Math.Round((millingToolBuilder.TlHeightBuilder.Value +
                                millingToolBuilder.TaperedShankLengthBuilder.Value) -
                                millingToolBuilder.TlZMountBuilder.Value)
                                / millingToolBuilder.TlDiameterBuilder.Value, 2);
                            break;

                    case Tool.Subtypes.DrillSpotFace:
                            millingToolBuilder = workPart.CAMSetup.CAMGroupCollection.CreateDrillSpotfaceToolBuilder(_tool);
                            _le.holderName = millingToolBuilder.TlHolderDescription;
                            _le.toolDiameter = string.Format("{0:0.0}", millingToolBuilder.TlDiameterBuilder.Value);
                            _le.toolZCount = millingToolBuilder.TlNumFlutesBuilder.Value.ToString();
                            _le.toolOffset = string.Format("{0:0.0}", (millingToolBuilder.TlHeightBuilder.Value + millingToolBuilder.TaperedShankLengthBuilder.Value) - millingToolBuilder.TlZMountBuilder.Value);
                            _le.diamToOffset = string.Format("{0:0.0}", Math.Round((millingToolBuilder.TlHeightBuilder.Value +
                                millingToolBuilder.TaperedShankLengthBuilder.Value) -
                                millingToolBuilder.TlZMountBuilder.Value)
                                / millingToolBuilder.TlDiameterBuilder.Value, 2);
                            break;

                    case Tool.Subtypes.DrillSpotDrill:
                            millingToolBuilder = workPart.CAMSetup.CAMGroupCollection.CreateDrillSpotdrillToolBuilder(_tool);
                            _le.holderName = millingToolBuilder.TlHolderDescription;
                            _le.toolDiameter = string.Format("{0:0.0}", millingToolBuilder.TlDiameterBuilder.Value);
                            _le.toolZCount = millingToolBuilder.TlNumFlutesBuilder.Value.ToString();
                            _le.toolOffset = string.Format("{0:0.0}", (millingToolBuilder.TlHeightBuilder.Value + millingToolBuilder.TaperedShankLengthBuilder.Value) - millingToolBuilder.TlZMountBuilder.Value);
                            _le.diamToOffset = string.Format("{0:0.0}", Math.Round((millingToolBuilder.TlHeightBuilder.Value +
                                millingToolBuilder.TaperedShankLengthBuilder.Value) -
                                millingToolBuilder.TlZMountBuilder.Value)
                                / millingToolBuilder.TlDiameterBuilder.Value, 2);
                            break;

                    case Tool.Subtypes.DrillBore:
                            millingToolBuilder = workPart.CAMSetup.CAMGroupCollection.CreateDrillBoreToolBuilder(_tool);
                            _le.holderName = millingToolBuilder.TlHolderDescription;
                            _le.toolDiameter = string.Format("{0:0.0}", millingToolBuilder.TlDiameterBuilder.Value);
                            _le.toolZCount = millingToolBuilder.TlNumFlutesBuilder.Value.ToString();
                            _le.toolOffset = string.Format("{0:0.0}", (millingToolBuilder.TlHeightBuilder.Value + millingToolBuilder.TaperedShankLengthBuilder.Value) - millingToolBuilder.TlZMountBuilder.Value);
                            _le.diamToOffset = string.Format("{0:0.0}", Math.Round((millingToolBuilder.TlHeightBuilder.Value +
                                millingToolBuilder.TaperedShankLengthBuilder.Value) -
                                millingToolBuilder.TlZMountBuilder.Value)
                                / millingToolBuilder.TlDiameterBuilder.Value, 2);
                            break;

                    case Tool.Subtypes.DrillReam:
                            millingToolBuilder = workPart.CAMSetup.CAMGroupCollection.CreateDrillReamerToolBuilder(_tool);
                            _le.holderName = millingToolBuilder.TlHolderDescription;
                            _le.toolDiameter = string.Format("{0:0.0}", millingToolBuilder.TlDiameterBuilder.Value);
                            _le.toolZCount = millingToolBuilder.TlNumFlutesBuilder.Value.ToString();
                            _le.toolOffset = string.Format("{0:0.0}", (millingToolBuilder.TlHeightBuilder.Value + millingToolBuilder.TaperedShankLengthBuilder.Value) - millingToolBuilder.TlZMountBuilder.Value);
                            _le.diamToOffset = string.Format("{0:0.0}", Math.Round((millingToolBuilder.TlHeightBuilder.Value +
                                millingToolBuilder.TaperedShankLengthBuilder.Value) -
                                millingToolBuilder.TlZMountBuilder.Value)
                                / millingToolBuilder.TlDiameterBuilder.Value, 2);
                            break;

                    case Tool.Subtypes.DrillCounterbore:
                            millingToolBuilder = workPart.CAMSetup.CAMGroupCollection.CreateDrillCounterboreToolBuilder(_tool);
                            _le.holderName = millingToolBuilder.TlHolderDescription;
                            _le.toolDiameter = string.Format("{0:0.0}", millingToolBuilder.TlDiameterBuilder.Value);
                            _le.toolZCount = millingToolBuilder.TlNumFlutesBuilder.Value.ToString();
                            _le.toolOffset = string.Format("{0:0.0}", (millingToolBuilder.TlHeightBuilder.Value + millingToolBuilder.TaperedShankLengthBuilder.Value) - millingToolBuilder.TlZMountBuilder.Value);
                            _le.diamToOffset = string.Format("{0:0.0}", Math.Round((millingToolBuilder.TlHeightBuilder.Value +
                                millingToolBuilder.TaperedShankLengthBuilder.Value) -
                                millingToolBuilder.TlZMountBuilder.Value)
                                / millingToolBuilder.TlDiameterBuilder.Value, 2);
                            break;

                    case Tool.Subtypes.DrillTap:
                            millingToolBuilder = workPart.CAMSetup.CAMGroupCollection.CreateDrillTapToolBuilder(_tool);
                            _le.holderName = millingToolBuilder.TlHolderDescription;
                            _le.toolDiameter = string.Format("{0:0.0}", millingToolBuilder.TlDiameterBuilder.Value);
                            _le.toolZCount = millingToolBuilder.TlNumFlutesBuilder.Value.ToString();
                            _le.toolOffset = string.Format("{0:0.0}", (millingToolBuilder.TlHeightBuilder.Value + millingToolBuilder.TaperedShankLengthBuilder.Value) - millingToolBuilder.TlZMountBuilder.Value);
                            _le.diamToOffset = string.Format("{0:0.0}", Math.Round((millingToolBuilder.TlHeightBuilder.Value +
                                millingToolBuilder.TaperedShankLengthBuilder.Value) -
                                millingToolBuilder.TlZMountBuilder.Value)
                                / millingToolBuilder.TlDiameterBuilder.Value, 2);
                            break;

                    case Tool.Subtypes.DrillThreadMill:
                            millingToolBuilder = workPart.CAMSetup.CAMGroupCollection.CreateDrillThreadMillToolBuilder(_tool);
                            _le.holderName = millingToolBuilder.TlHolderDescription;
                            _le.toolDiameter = string.Format("{0:0.0}", millingToolBuilder.TlDiameterBuilder.Value);
                            _le.toolZCount = millingToolBuilder.TlNumFlutesBuilder.Value.ToString();
                            _le.toolOffset = string.Format("{0:0.0}", (millingToolBuilder.TlHeightBuilder.Value + millingToolBuilder.TaperedShankLengthBuilder.Value) - millingToolBuilder.TlZMountBuilder.Value);
                            _le.diamToOffset = string.Format("{0:0.0}", Math.Round((millingToolBuilder.TlHeightBuilder.Value +
                                millingToolBuilder.TaperedShankLengthBuilder.Value) -
                                millingToolBuilder.TlZMountBuilder.Value)
                                / millingToolBuilder.TlDiameterBuilder.Value, 2);
                            break;

                    case Tool.Subtypes.DrillStep:
                            millingToolBuilder = workPart.CAMSetup.CAMGroupCollection.CreateDrillStepToolBuilder(_tool);
                            _le.holderName = millingToolBuilder.TlHolderDescription;
                            _le.toolDiameter = string.Format("{0:0.0}", millingToolBuilder.TlDiameterBuilder.Value);
                            _le.toolZCount = millingToolBuilder.TlNumFlutesBuilder.Value.ToString();
                            _le.toolOffset = string.Format("{0:0.0}", (millingToolBuilder.TlHeightBuilder.Value + millingToolBuilder.TaperedShankLengthBuilder.Value) - millingToolBuilder.TlZMountBuilder.Value);
                            _le.diamToOffset = string.Format("{0:0.0}", Math.Round((millingToolBuilder.TlHeightBuilder.Value +
                                millingToolBuilder.TaperedShankLengthBuilder.Value) -
                                millingToolBuilder.TlZMountBuilder.Value)
                                / millingToolBuilder.TlDiameterBuilder.Value, 2);
                            break;
                    }  
            break;
            }             
        }

        public void GetOpParameters(Operation op, LibraryElement _leOp, Part workPart)
        {            
            int OperationType;
            ufSession.Oper.AskOperType(op.Tag, out OperationType);            

            if (OperationType == 110)
            {
                var _operB = workPart.CAMSetup.CAMOperationCollection.CreatePlanarMillingBuilder(op);

                var fpt3 = Math.Round(_operB.FeedsBuilder.FeedPerToothBuilder.Value, 3);
                var fcb = _operB.FeedsBuilder.FeedCutBuilder.Value.ToString();
                var srb = _operB.FeedsBuilder.SpindleRpmBuilder.Value.ToString();
                var ssb = _operB.FeedsBuilder.SurfaceSpeedBuilder.Value;

                _leOp.fZ = string.Format("{0:0.0}", fpt3);
                _leOp.Vc = string.Format("{0:0.0}", ssb);
                _leOp.CutStep = string.Format("{0:0.0}", _operB.DepthPerCut.Value);
                _leOp.CutWidth = string.Format("{0:0.0}", _operB.BndStepover.PercentToolFlatBuilder.Value);
                _leOp.operationType = OperationType.ToString();
            }
            else if (OperationType == 260)
            {
                var _operB = workPart.CAMSetup.CAMOperationCollection.CreateCavityMillingBuilder(op);
                var fcb = _operB.FeedsBuilder.FeedCutBuilder.Value.ToString();
                var srb = _operB.FeedsBuilder.SpindleRpmBuilder.Value.ToString();
                var ssb = _operB.FeedsBuilder.SurfaceSpeedBuilder.Value;
                var fpt3 = Math.Round(_operB.FeedsBuilder.FeedPerToothBuilder.Value, 3);
                _leOp.fZ = string.Format("{0:0.0}", fpt3);
                _leOp.Vc = string.Format("{0:0.0}", ssb);
                _leOp.CutStep = string.Format("{0:0.0}", _operB.DepthPerCut.Value);
                _leOp.CutWidth = string.Format("{0:0.0}", _operB.BndStepover.PercentToolFlatBuilder.Value);
                _leOp.operationType = OperationType.ToString();
            }
            else if (OperationType == 261)
            {
                var _operB = workPart.CAMSetup.CAMOperationCollection.CreateFaceMillingBuilder(op);
                var fpt3 = Math.Round(_operB.FeedsBuilder.FeedPerToothBuilder.Value, 3);
                var fpt = fpt3.ToString();
                var fcb = _operB.FeedsBuilder.FeedCutBuilder.Value.ToString();
                var srb = _operB.FeedsBuilder.SpindleRpmBuilder.Value.ToString();
                var ssb = _operB.FeedsBuilder.SurfaceSpeedBuilder.Value;

                _leOp.fZ = string.Format("{0:0.0}", fpt3);
                _leOp.Vc = string.Format("{0:0.0}", ssb);
                _leOp.CutStep = string.Format("{0:0.0}", _operB.DepthPerCut.Value);
                _leOp.CutWidth = string.Format("{0:0.0}", _operB.BndStepover.PercentToolFlatBuilder.Value);
                _leOp.operationType = OperationType.ToString();
            }
            else if (OperationType == 3100)
            {
                var _operB = workPart.CAMSetup.CAMOperationCollection.CreateGrooveMillingBuilder(op);
                var fpt3 = Math.Round(_operB.FeedsBuilder.FeedPerToothBuilder.Value, 3);
                var fpt = fpt3.ToString();
                var fcb = _operB.FeedsBuilder.FeedCutBuilder.Value.ToString();
                var srb = _operB.FeedsBuilder.SpindleRpmBuilder.Value.ToString();
                var ssb = _operB.FeedsBuilder.SurfaceSpeedBuilder.Value;

                _leOp.fZ = string.Format("{0:0.0}", fpt3);
                _leOp.Vc = string.Format("{0:0.0}", ssb);
                _leOp.CutStep = string.Format("{0:0.0}", _operB.DepthPerCut.Value);
                _leOp.CutWidth = string.Format("{0:0.0}", _operB.BndStepover.PercentToolFlatBuilder.Value);
                _leOp.operationType = OperationType.ToString();
            }
            else if (OperationType == 3200)
            {
                var _operB = workPart.CAMSetup.CAMOperationCollection.CreateChamferMillingBuilder(op);
                var fpt3 = Math.Round(_operB.FeedsBuilder.FeedPerToothBuilder.Value, 3);
                var fpt = fpt3.ToString();
                var fcb = _operB.FeedsBuilder.FeedCutBuilder.Value.ToString();
                var srb = _operB.FeedsBuilder.SpindleRpmBuilder.Value.ToString();
                var ssb = _operB.FeedsBuilder.SurfaceSpeedBuilder.Value;

                _leOp.fZ = string.Format("{0:0.0}", fpt3);
                _leOp.Vc = string.Format("{0:0.0}", ssb);
                _leOp.CutStep = string.Format("{0:0.0}", 0.0);
                _leOp.CutWidth = string.Format("{0:0.0}", 0.0);
                _leOp.operationType = OperationType.ToString();
                               
            }
            /*
            else if (OperationType == 3800)
            {
                NXOpen.CAM.CAMObject[] p = new CAMObject[1];
                p[0] = op;
                NXOpen.CAM.ObjectsFeedsBuilder _operB = workPart.CAMSetup.CreateFeedsBuilder(p);
                var _operC = workPart.CAMSetup.CAMOperationCollection.CreateMultiAxisDeburringBuilder(op);
                var fpt3 = Math.Round(_operB.FeedsBuilder.FeedPerToothBuilder.Value, 3);

                var fpt = fpt3.ToString();
                var fcb = _operB.FeedsBuilder.FeedCutBuilder.Value.ToString();
                var srb = _operB.FeedsBuilder.SpindleRpmBuilder.Value.ToString();
                var ssb = _operB.FeedsBuilder.SurfaceSpeedBuilder.Value.ToString();

                var wall = 0.0;
                var flr = 0.0;

                var opCorr = _operC.ToolChangeSetting.CutcomRegister.Value;
                //var cutCom = _operC.NonCuttingBuilder.CutcomOutputContactPoint; //true/False
                var cutComPlane = NcmPlanarBuilder.CutcomTypes.None; //by type
                var opAdj = _operC.ToolChangeSetting.AdjustRegister.Value; // length adjust register                 
                //var cutCom = _operB.NonCuttingBuilder.CutcomOutputContactPoint; //true/False
            }
            */
            else if (OperationType == 262)
            {
                var _operB = workPart.CAMSetup.CAMOperationCollection.CreateVolumeBased25dMillingOperationBuilder(op);
                var fpt3 = Math.Round(_operB.FeedsBuilder.FeedPerToothBuilder.Value, 3);
                var fpt = fpt3.ToString();
                var fcb = _operB.FeedsBuilder.FeedCutBuilder.Value.ToString();
                var srb = _operB.FeedsBuilder.SpindleRpmBuilder.Value.ToString();
                var ssb = _operB.FeedsBuilder.SurfaceSpeedBuilder.Value;

                _leOp.fZ = string.Format("{0:0.0}", fpt3);
                _leOp.Vc = string.Format("{0:0.0}", ssb);
                _leOp.CutStep = string.Format("{0:0.0}", _operB.DepthPerCut.Value);
                _leOp.CutWidth = string.Format("{0:0.0}", _operB.BndStepover.PercentToolFlatBuilder.Value);
                _leOp.operationType = OperationType.ToString();
            }
            else if (OperationType == 2700)
            {
                var _operB = workPart.CAMSetup.CAMOperationCollection.CreateCylinderMillingBuilder(op);
                var fpt3 = Math.Round(_operB.FeedsBuilder.FeedPerToothBuilder.Value, 3);
                var fpt = fpt3.ToString();
                var fcb = _operB.FeedsBuilder.FeedCutBuilder.Value.ToString();
                var srb = _operB.FeedsBuilder.SpindleRpmBuilder.Value.ToString();
                var ssb = _operB.FeedsBuilder.SurfaceSpeedBuilder.Value;

                _leOp.fZ = string.Format("{0:0.0}", fpt3);
                _leOp.Vc = string.Format("{0:0.0}", ssb);
                _leOp.CutStep = string.Format("{0:0.0}", 0.0);
                _leOp.CutWidth = string.Format("{0:0.0}", 0.0);
                _leOp.operationType = OperationType.ToString();
            }
            else if (OperationType == 1700)
            {
                var _operB = workPart.CAMSetup.CAMOperationCollection.CreateThreadMillingBuilder(op);
                var fpt3 = Math.Round(_operB.FeedsBuilder.FeedPerToothBuilder.Value, 3);
                var fpt = fpt3.ToString();
                var fcb = _operB.FeedsBuilder.FeedCutBuilder.Value.ToString();
                var srb = _operB.FeedsBuilder.SpindleRpmBuilder.Value.ToString();
                var ssb = _operB.FeedsBuilder.SurfaceSpeedBuilder.Value;

                _leOp.fZ = string.Format("{0:0.0}", fpt3);
                _leOp.Vc = string.Format("{0:0.0}", ssb);
                _leOp.CutStep = string.Format("{0:0.0}", 0.0);
                _leOp.CutWidth = string.Format("{0:0.0}", 0.0);
                _leOp.operationType = OperationType.ToString();
            }
            else if (OperationType == 265)
            {
                var _operB = workPart.CAMSetup.CAMOperationCollection.CreatePlungeMillingBuilder(op);
                var fpt3 = Math.Round(_operB.FeedsBuilder.FeedPerToothBuilder.Value, 3);
                var fpt = fpt3.ToString();
                var fcb = _operB.FeedsBuilder.FeedCutBuilder.Value.ToString();
                var srb = _operB.FeedsBuilder.SpindleRpmBuilder.Value.ToString();
                var ssb = _operB.FeedsBuilder.SurfaceSpeedBuilder.Value;

                _leOp.fZ = string.Format("{0:0.0}", fpt3);
                _leOp.Vc = string.Format("{0:0.0}", ssb);
                _leOp.CutStep = string.Format("{0:0.0}", _operB.DepthPerCut.Value);
                _leOp.CutWidth = string.Format("{0:0.0}", _operB.BndStepover.PercentToolFlatBuilder.Value);
                _leOp.operationType = OperationType.ToString();
            }
            else if (OperationType.ToString() == "Engraving")
            {
                var _operB = workPart.CAMSetup.CAMOperationCollection.CreateEngravingBuilder(op);
                var fpt3 = Math.Round(_operB.FeedsBuilder.FeedPerToothBuilder.Value, 3);
                var fpt = fpt3.ToString();
                var fcb = _operB.FeedsBuilder.FeedCutBuilder.Value.ToString();
                var srb = _operB.FeedsBuilder.SpindleRpmBuilder.Value.ToString();
                var ssb = _operB.FeedsBuilder.SurfaceSpeedBuilder.Value;

                _leOp.fZ = string.Format("{0:0.0}", fpt3);
                _leOp.Vc = string.Format("{0:0.0}", ssb);
                _leOp.CutStep = string.Format("{0:0.0}", 0.0);
                _leOp.CutWidth = string.Format("{0:0.0}", 0.0);
                _leOp.operationType = OperationType.ToString();

            }
            else if (OperationType == 1100)
            {
                var _operB = workPart.CAMSetup.CAMOperationCollection.CreateMillMachineControlBuilder(op);
                var fpt3 = Math.Round(_operB.FeedsBuilder.FeedPerToothBuilder.Value, 3);
                var fpt = fpt3.ToString();
                var fcb = _operB.FeedsBuilder.FeedCutBuilder.Value.ToString();
                var srb = _operB.FeedsBuilder.SpindleRpmBuilder.Value.ToString();
                var ssb = _operB.FeedsBuilder.SurfaceSpeedBuilder.Value;
                var opAdj = _operB.ToolChangeSetting.AdjustRegister.Value;
                var opCorr = _operB.ToolChangeSetting.CutcomRegister.Value;

                _leOp.fZ = string.Format("{0:0.0}", fpt3);
                _leOp.Vc = string.Format("{0:0.0}", ssb);
                _leOp.CutStep = string.Format("{0:0.0}", 0.0);
                _leOp.CutWidth = string.Format("{0:0.0}", 0.0);
                _leOp.operationType = OperationType.ToString();
            }
            else if (OperationType == 263)
            {
                var _operB = workPart.CAMSetup.CAMOperationCollection.CreateZlevelMillingBuilder(op);
                var fpt3 = Math.Round(_operB.FeedsBuilder.FeedPerToothBuilder.Value, 3);
                var fcb = _operB.FeedsBuilder.FeedCutBuilder.Value.ToString();
                var srb = _operB.FeedsBuilder.SpindleRpmBuilder.Value.ToString();
                var ssb = _operB.FeedsBuilder.SurfaceSpeedBuilder.Value;

                _leOp.fZ = string.Format("{0:0.0}", fpt3);
                _leOp.Vc = string.Format("{0:0.0}", ssb);
                _leOp.CutStep = string.Format("{0:0.0}", _operB.DepthPerCut.Value);
                _leOp.CutWidth = string.Format("{0:0.0}", _operB.BndStepover.PercentToolFlatBuilder.Value);
                _leOp.operationType = OperationType.ToString();
            }
            else if (OperationType == 210 || OperationType == 211)
            {
                var _operB = workPart.CAMSetup.CAMOperationCollection.CreateSurfaceContourBuilder(op);
                var fpt3 = Math.Round(_operB.FeedsBuilder.FeedPerToothBuilder.Value, 3);
                var fpt = fpt3.ToString();
                var fcb = _operB.FeedsBuilder.FeedCutBuilder.Value.ToString();
                var srb = _operB.FeedsBuilder.SpindleRpmBuilder.Value.ToString();
                var ssb = _operB.FeedsBuilder.SurfaceSpeedBuilder.Value;

                _leOp.fZ = string.Format("{0:0.0}", fpt3);
                _leOp.Vc = string.Format("{0:0.0}", ssb);
                _leOp.CutStep = string.Format("{0:0.0}", 0.0);
                _leOp.CutWidth = string.Format("{0:0.0}", 0.0);
                _leOp.operationType = OperationType.ToString();
            }
            else if (OperationType == 900)
            {
                var _operB = workPart.CAMSetup.CAMOperationCollection.CreateGmcopBuilder(op);
                var fpt3 = Math.Round(_operB.FeedsBuilder.FeedPerToothBuilder.Value, 3);
                var fpt = fpt3.ToString();
                var fcb = _operB.FeedsBuilder.FeedCutBuilder.Value.ToString();
                var srb = _operB.FeedsBuilder.SpindleRpmBuilder.Value.ToString();
                var ssb = _operB.FeedsBuilder.SurfaceSpeedBuilder.Value;

                _leOp.fZ = string.Format("{0:0.0}", fpt3);
                _leOp.Vc = string.Format("{0:0.0}", ssb);
                _leOp.CutStep = string.Format("{0:0.0}", 0.0);
                _leOp.CutWidth = string.Format("{0:0.0}", 0.0);
                _leOp.operationType = OperationType.ToString();
            }
            else if (OperationType == 3000)
            {
                var _operB = workPart.CAMSetup.CAMOperationCollection.CreateHoleDrillingBuilder(op);
                var fpt3 = Math.Round(_operB.FeedsBuilder.FeedPerToothBuilder.Value, 3);
                var fpt = fpt3.ToString();
                var fcb = _operB.FeedsBuilder.FeedCutBuilder.Value.ToString();
                var srb = _operB.FeedsBuilder.SpindleRpmBuilder.Value.ToString();
                var ssb = _operB.FeedsBuilder.SurfaceSpeedBuilder.Value;

                _leOp.fZ = string.Format("{0:0.0}", fpt3);
                _leOp.Vc = string.Format("{0:0.0}", ssb);
                _leOp.CutStep = string.Format("{0:0.0}", 0.0);
                _leOp.CutWidth = string.Format("{0:0.0}", 0.0);
                _leOp.operationType = OperationType.ToString();
            }
            else if (OperationType == 3300)
            {
                var _operB = workPart.CAMSetup.CAMOperationCollection.CreateRadialGrooveMillingBuilder(op);
                var fpt = _operB.FeedsBuilder.FeedPerToothBuilder.Value;
                var fcb = _operB.FeedsBuilder.FeedCutBuilder.Value.ToString();
                var srb = _operB.FeedsBuilder.SpindleRpmBuilder.Value.ToString();
                var ssb = _operB.FeedsBuilder.SurfaceSpeedBuilder.Value;

                _leOp.fZ = string.Format("{0:0.0}", fpt);
                _leOp.Vc = string.Format("{0:0.0}", ssb);
                _leOp.CutStep = string.Format("{0:0.0}", 0.0);
                _leOp.CutWidth = string.Format("{0:0.0}", 0.0);
                _leOp.operationType = OperationType.ToString();
            }
            else if (OperationType == 600)
            {
                var _operB = workPart.CAMSetup.CAMOperationCollection.CreateHoleMakingBuilder(op);
                var fpt3 = Math.Round(_operB.FeedsBuilder.FeedPerToothBuilder.Value, 3);
                var fpt = fpt3.ToString();
                var fcb = _operB.FeedsBuilder.FeedCutBuilder.Value.ToString();
                var srb = _operB.FeedsBuilder.SpindleRpmBuilder.Value.ToString();
                var ssb = _operB.FeedsBuilder.SurfaceSpeedBuilder.Value;

                _leOp.fZ = string.Format("{0:0.0}", fpt3);
                _leOp.Vc = string.Format("{0:0.0}", ssb);
                _leOp.CutStep = string.Format("{0:0.0}", 0.0);
                _leOp.CutWidth = string.Format("{0:0.0}", 0.0);
                _leOp.operationType = OperationType.ToString();
            }
            /*
            else if (optN.Contains("FeatureMilling"))
            {
                var _operB = workPart.CAMSetup.CAMOperationCollection.CreateFeatureMillingBuilder(op);
                var fpt3 = Math.Round(_operB.FeedsBuilder.FeedPerToothBuilder.Value, 3);
                var fpt = fpt3.ToString();
                var fcb = _operB.FeedsBuilder.FeedCutBuilder.Value.ToString();
                var srb = _operB.FeedsBuilder.SpindleRpmBuilder.Value.ToString();
                var ssb = _operB.FeedsBuilder.SurfaceSpeedBuilder.Value.ToString();
            }
            */
            else if (OperationType == 450)
            {
                var _operB = workPart.CAMSetup.CAMOperationCollection.CreatePointToPointBuilder(op);
                var fpt3 = Math.Round(_operB.FeedsBuilder.FeedPerToothBuilder.Value, 3);
                var fpt = fpt3.ToString();
                var fcb = _operB.FeedsBuilder.FeedCutBuilder.Value.ToString();
                var srb = _operB.FeedsBuilder.SpindleRpmBuilder.Value.ToString();
                var ssb = _operB.FeedsBuilder.SurfaceSpeedBuilder.Value;

                _leOp.fZ = string.Format("{0:0.0}", fpt3);
                _leOp.Vc = string.Format("{0:0.0}", ssb);
                _leOp.CutStep = string.Format("{0:0.0}", 0.0);
                _leOp.CutWidth = string.Format("{0:0.0}", 0.0);
                _leOp.operationType = OperationType.ToString();
            }
            else if (OperationType == 530)
            {
                var _operB = workPart.CAMSetup.CAMOperationCollection.CreateTeachmodeTurningBuilder(op);
                var fpt3 = Math.Round(_operB.FeedsBuilder.FeedPerToothBuilder.Value, 3);
                var fpt = fpt3.ToString();
                var fcb = _operB.FeedsBuilder.FeedCutBuilder.Value.ToString();
                var srb = _operB.FeedsBuilder.SpindleRpmBuilder.Value.ToString();
                var ssb = _operB.FeedsBuilder.SurfaceSpeedBuilder.Value;

                _leOp.fZ = string.Format("{0:0.0}", fpt3);
                _leOp.Vc = string.Format("{0:0.0}", ssb);
                _leOp.CutStep = string.Format("{0:0.0}", 0.0);
                _leOp.CutWidth = string.Format("{0:0.0}", 0.0);
                _leOp.operationType = OperationType.ToString();
            }
            else if (OperationType == 550)
            {
                var _operB = workPart.CAMSetup.CAMOperationCollection.CreateCenterlineDrillTurningBuilder(op);
                var fpt3 = Math.Round(_operB.FeedsBuilder.FeedPerToothBuilder.Value, 3);
                var fpt = fpt3.ToString();
                var fcb = _operB.FeedsBuilder.FeedCutBuilder.Value.ToString();
                var srb = _operB.FeedsBuilder.SpindleRpmBuilder.Value.ToString();
                var ssb = _operB.FeedsBuilder.SurfaceSpeedBuilder.Value;

                _leOp.fZ = string.Format("{0:0.0}", fpt3);
                _leOp.Vc = string.Format("{0:0.0}", ssb);
                _leOp.CutStep = string.Format("{0:0.0}", 0.0);
                _leOp.CutWidth = string.Format("{0:0.0}", 0.0);
                _leOp.operationType = OperationType.ToString();
            }
            else if (OperationType == 510)
            {
                var _operB = workPart.CAMSetup.CAMOperationCollection.CreateRoughTurningBuilder(op);
                var fpt3 = Math.Round(_operB.FeedsBuilder.FeedPerToothBuilder.Value, 3);
                var fpt = fpt3.ToString();
                var fcb = _operB.FeedsBuilder.FeedCutBuilder.Value.ToString();
                var srb = _operB.FeedsBuilder.SpindleRpmBuilder.Value.ToString();
                var ssb = _operB.FeedsBuilder.SurfaceSpeedBuilder.Value;

                _leOp.fZ = string.Format("{0:0.0}", fpt3);
                _leOp.Vc = string.Format("{0:0.0}", ssb);
                _leOp.CutStep = string.Format("{0:0.0}", 0.0);
                _leOp.CutWidth = string.Format("{0:0.0}", 0.0);
                _leOp.operationType = OperationType.ToString();
            }
            else if (OperationType == 520)
            {
                var _operB = workPart.CAMSetup.CAMOperationCollection.CreateFinishTurningBuilder(op);
                var fpt3 = Math.Round(_operB.FeedsBuilder.FeedPerToothBuilder.Value, 3);
                var fpt = fpt3.ToString();
                var fcb = _operB.FeedsBuilder.FeedCutBuilder.Value.ToString();
                var srb = _operB.FeedsBuilder.SpindleRpmBuilder.Value.ToString();
                var ssb = _operB.FeedsBuilder.SurfaceSpeedBuilder.Value;

                _leOp.fZ = string.Format("{0:0.0}", fpt3);
                _leOp.Vc = string.Format("{0:0.0}", ssb);
                _leOp.CutStep = string.Format("{0:0.0}", 0.0);
                _leOp.CutWidth = string.Format("{0:0.0}", 0.0);
                _leOp.operationType = OperationType.ToString();
            }
            else if (OperationType == 540)
            {
                var _operB = workPart.CAMSetup.CAMOperationCollection.CreateThreadTurningBuilder(op);
                var fpt3 = Math.Round(_operB.FeedsBuilder.FeedPerToothBuilder.Value, 3);
                var fpt = fpt3.ToString();
                var fcb = _operB.FeedsBuilder.FeedCutBuilder.Value.ToString();
                var srb = _operB.FeedsBuilder.SpindleRpmBuilder.Value.ToString();
                var ssb = _operB.FeedsBuilder.SurfaceSpeedBuilder.Value;

                _leOp.fZ = string.Format("{0:0.0}", fpt3);
                _leOp.Vc = string.Format("{0:0.0}", ssb);
                _leOp.CutStep = string.Format("{0:0.0}", 0.0);
                _leOp.CutWidth = string.Format("{0:0.0}", 0.0);
                _leOp.operationType = OperationType.ToString();
            }
            else if (OperationType == 1200)
            {
                var _operB = workPart.CAMSetup.CAMOperationCollection.CreateLatheMachineControlBuilder(op);
                var fpt3 = Math.Round(_operB.FeedsBuilder.FeedPerToothBuilder.Value, 3);
                var fpt = fpt3.ToString();
                var fcb = _operB.FeedsBuilder.FeedCutBuilder.Value.ToString();
                var srb = _operB.FeedsBuilder.SpindleRpmBuilder.Value.ToString();
                var ssb = _operB.FeedsBuilder.SurfaceSpeedBuilder.Value;

                _leOp.fZ = string.Format("{0:0.0}", fpt3);
                _leOp.Vc = string.Format("{0:0.0}", ssb);
                _leOp.CutStep = string.Format("{0:0.0}", 0.0);
                _leOp.CutWidth = string.Format("{0:0.0}", 0.0);
                _leOp.operationType = OperationType.ToString();
            } 
        }        
    }
}
