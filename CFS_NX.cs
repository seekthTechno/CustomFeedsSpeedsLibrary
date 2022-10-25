using System;
using NXOpen;
using NXOpen.BlockStyler;

namespace CustomFeedsSpeedsLibrary
{
    public class CFS_NX
    {
        //class members
        public DataService _program;
        public static Session theSession = null;
        public static UI theUI = null;
        public string theDlxFileName;
        public NXOpen.BlockStyler.BlockDialog theDialog;
        public NXOpen.BlockStyler.Group selectionGroup;// Block type: Group
        public NXOpen.BlockStyler.SelectObject opsOrGroupSelection;// Block type: Selection
        public NXOpen.BlockStyler.Label labelObject;// Block type: Label
        public NXOpen.BlockStyler.Group machineGroup1;// Block type: Group
        public NXOpen.BlockStyler.Enumeration machineGroup;// Block type: Enumeration
        public NXOpen.BlockStyler.Enumeration machine;// Block type: Enumeration
        public NXOpen.BlockStyler.Group materialGroup1;// Block type: Group
        public NXOpen.BlockStyler.DoubleBlock Rockwell;// Block type: Double
        public NXOpen.BlockStyler.Enumeration materialList;// Block type: Enumeration
        public NXOpen.BlockStyler.Group workpieceGroup1;// Block type: Group
        public NXOpen.BlockStyler.Label wrkpcLabel;// Block type: Label
        public NXOpen.BlockStyler.IntegerBlock integer0;// Block type: Integer
        public NXOpen.BlockStyler.Button applyFNS;// Block type: Button
        public NXOpen.BlockStyler.Button saveFNS;// Block type: Button
       
        //------------------------------------------------------------------------------
        //Bit Option for Property: SnapPointTypesEnabled
        //------------------------------------------------------------------------------
        public static readonly int SnapPointTypesEnabled_UserDefined = (1 << 0);
        public static readonly int SnapPointTypesEnabled_Inferred = (1 << 1);
        public static readonly int SnapPointTypesEnabled_ScreenPosition = (1 << 2);
        public static readonly int SnapPointTypesEnabled_EndPoint = (1 << 3);
        public static readonly int SnapPointTypesEnabled_MidPoint = (1 << 4);
        public static readonly int SnapPointTypesEnabled_ControlPoint = (1 << 5);
        public static readonly int SnapPointTypesEnabled_Intersection = (1 << 6);
        public static readonly int SnapPointTypesEnabled_ArcCenter = (1 << 7);
        public static readonly int SnapPointTypesEnabled_QuadrantPoint = (1 << 8);
        public static readonly int SnapPointTypesEnabled_ExistingPoint = (1 << 9);
        public static readonly int SnapPointTypesEnabled_PointonCurve = (1 << 10);
        public static readonly int SnapPointTypesEnabled_PointonSurface = (1 << 11);
        public static readonly int SnapPointTypesEnabled_PointConstructor = (1 << 12);
        public static readonly int SnapPointTypesEnabled_TwocurveIntersection = (1 << 13);
        public static readonly int SnapPointTypesEnabled_TangentPoint = (1 << 14);
        public static readonly int SnapPointTypesEnabled_Poles = (1 << 15);
        public static readonly int SnapPointTypesEnabled_BoundedGridPoint = (1 << 16);
        public static readonly int SnapPointTypesEnabled_FacetVertexPoint = (1 << 17);
        //------------------------------------------------------------------------------
        //Bit Option for Property: SnapPointTypesOnByDefault
        //------------------------------------------------------------------------------
        public static readonly int SnapPointTypesOnByDefault_EndPoint = (1 << 3);
        public static readonly int SnapPointTypesOnByDefault_MidPoint = (1 << 4);
        public static readonly int SnapPointTypesOnByDefault_ControlPoint = (1 << 5);
        public static readonly int SnapPointTypesOnByDefault_Intersection = (1 << 6);
        public static readonly int SnapPointTypesOnByDefault_ArcCenter = (1 << 7);
        public static readonly int SnapPointTypesOnByDefault_QuadrantPoint = (1 << 8);
        public static readonly int SnapPointTypesOnByDefault_ExistingPoint = (1 << 9);
        public static readonly int SnapPointTypesOnByDefault_PointonCurve = (1 << 10);
        public static readonly int SnapPointTypesOnByDefault_PointonSurface = (1 << 11);
        public static readonly int SnapPointTypesOnByDefault_PointConstructor = (1 << 12);
        public static readonly int SnapPointTypesOnByDefault_BoundedGridPoint = (1 << 16);
        
        public CFS_NX()
        {
            try
            {
                theSession = Session.GetSession();
                theUI = UI.GetUI();
                theDlxFileName = "CFS_NX.dlx";
                theDialog = theUI.CreateDialog(theDlxFileName);
                theDialog.AddUpdateHandler(new NXOpen.BlockStyler.BlockDialog.Update(update_cb));
                theDialog.AddInitializeHandler(new NXOpen.BlockStyler.BlockDialog.Initialize(initialize_cb));
                theDialog.AddFocusNotifyHandler(new NXOpen.BlockStyler.BlockDialog.FocusNotify(focusNotify_cb));
                theDialog.AddKeyboardFocusNotifyHandler(new NXOpen.BlockStyler.BlockDialog.KeyboardFocusNotify(keyboardFocusNotify_cb));
                theDialog.AddEnableOKButtonHandler(new NXOpen.BlockStyler.BlockDialog.EnableOKButton(enableOKButton_cb));
                theDialog.AddDialogShownHandler(new NXOpen.BlockStyler.BlockDialog.DialogShown(dialogShown_cb));                
            }
            catch (Exception ex)
            {
                //---- Enter your exception handling code here -----
                throw ex;
            }
        }

        public static void Main()
        {
            CFS_NX theCFS_NX = null;
            try
            {
                theCFS_NX = new CFS_NX();
                theCFS_NX._program = new DataService(theCFS_NX);
                theCFS_NX.Show();
            }
            catch (Exception ex)
            {
                //---- Enter your exception handling code here -----
                theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString());
            }
            finally
            {
                if (theCFS_NX != null)
                    theCFS_NX.Dispose();
                theCFS_NX = null;
            }
        }
        
     public static int GetUnloadOption(string arg)
    {        
         return System.Convert.ToInt32(Session.LibraryUnloadOption.Immediately);        
    }    
    
    public static void UnloadLibrary(string arg)
    {
        try
        {
            //---- Enter your code here -----
        }
        catch (Exception ex)
        {
            //---- Enter your exception handling code here -----
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString());
        }
    }

        public NXOpen.UIStyler.DialogResponse Show()
        {
            try
            {
                theDialog.Show();
            }
            catch (Exception ex)
            {
                //---- Enter your exception handling code here -----
                theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString());
            }
            return 0;
        }
        
        public void Dispose()
        {
            if (theDialog != null)
            {
                theDialog.Dispose();
                theDialog = null;
            }
        }

        public void initialize_cb()
        {
            try
            {
                selectionGroup = (NXOpen.BlockStyler.Group)theDialog.TopBlock.FindBlock("selectionGroup");
                opsOrGroupSelection = (NXOpen.BlockStyler.SelectObject)theDialog.TopBlock.FindBlock("opsOrGroupSelection");
                labelObject = (NXOpen.BlockStyler.Label)theDialog.TopBlock.FindBlock("labelObject");
                machineGroup1 = (NXOpen.BlockStyler.Group)theDialog.TopBlock.FindBlock("machineGroup1");
                machineGroup = (NXOpen.BlockStyler.Enumeration)theDialog.TopBlock.FindBlock("machineGroup");
                machine = (NXOpen.BlockStyler.Enumeration)theDialog.TopBlock.FindBlock("machine");
                materialGroup1 = (NXOpen.BlockStyler.Group)theDialog.TopBlock.FindBlock("materialGroup1");
                Rockwell = (NXOpen.BlockStyler.DoubleBlock)theDialog.TopBlock.FindBlock("Rockwell");
                materialList = (NXOpen.BlockStyler.Enumeration)theDialog.TopBlock.FindBlock("materialList");
                workpieceGroup1 = (NXOpen.BlockStyler.Group)theDialog.TopBlock.FindBlock("workpieceGroup1");
                wrkpcLabel = (NXOpen.BlockStyler.Label)theDialog.TopBlock.FindBlock("wrkpcLabel");
                integer0 = (NXOpen.BlockStyler.IntegerBlock)theDialog.TopBlock.FindBlock("integer0");
                applyFNS = (NXOpen.BlockStyler.Button)theDialog.TopBlock.FindBlock("applyFNS");
                saveFNS = (NXOpen.BlockStyler.Button)theDialog.TopBlock.FindBlock("saveFNS");
            }
            catch (Exception ex)
            {
                theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString());
            }
        }

        public void dialogShown_cb()
        {
            try
            {
                _program.Initialise_Data();
            }
            catch (Exception ex)
            {
                theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString());
            }
        }

        public int update_cb(NXOpen.BlockStyler.UIBlock block)
        {
            try
            {
                _program.updateCb(block);
                if (block == opsOrGroupSelection)
                {
                    //---------Enter your code here-----------
                }
                else if (block == labelObject)
                {
                    //---------Enter your code here-----------
                }
                else if (block == machineGroup)
                {
                    //---------Enter your code here-----------
                }
                else if (block == machine)
                {
                    //---------Enter your code here-----------
                }
                else if (block == Rockwell)
                {
                    //---------Enter your code here-----------
                }
                else if (block == materialList)
                {
                    //---------Enter your code here-----------
                }
                else if (block == wrkpcLabel)
                {
                    //---------Enter your code here-----------
                }
                else if (block == integer0)
                {
                    //---------Enter your code here-----------
                }
                else if (block == applyFNS)
                {
                    //---------Enter your code here-----------
                }
                else if (block == saveFNS)
                {
                    //---------Enter your code here-----------
                }
            }
            catch (Exception ex)
            {
                theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString());
            }
            return 0;
        }
        
        public void focusNotify_cb(NXOpen.BlockStyler.UIBlock block, bool focus)
        {
            try
            {
                
            }
            catch (Exception ex)
            {
                theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString());
            }
        }

        public void keyboardFocusNotify_cb(NXOpen.BlockStyler.UIBlock block, bool focus)
        {
            try
            {
                
            }
            catch (Exception ex)
            {
                theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString());
            }
        }

        public bool enableOKButton_cb()
        {
            try
            {
               
            }
            catch (Exception ex)
            {
                theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString());
            }
            return true;
        }

        public PropertyList GetBlockProperties(string blockID)
        {
            PropertyList plist = null;
            try
            {
                plist = theDialog.GetBlockProperties(blockID);
            }
            catch (Exception ex)
            {
                theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString());
            }
            return plist;
        }

    }
}
