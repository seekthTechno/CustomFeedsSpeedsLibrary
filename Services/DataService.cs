using System.Collections.Generic;
using System.Linq;
using System.Text;
using NXOpen;
using NXOpen.UF;
using NXOpen.BlockStyler;
using CustomFeedsSpeedsLibrary.Data;
using NXOpen.Utilities;
using NXOpen.CAM;

namespace CustomFeedsSpeedsLibrary
{
    public class DataService
    {
        private static Session _Session = Session.GetSession();
        private static UI _UI = UI.GetUI();        
        private CFS_NX _Dialog;
        private static UFSession Ufs = UFSession.GetUFSession();

        private SelectionCntrl _uiCtrl;
        private Attrbutes _attrbutes;        
        private Encoding Enc = Encoding.GetEncoding(1251);
        private DataService ds;

        public string machineName { get; set; }
        public string partMaterial { get; set; }
        public double partHardness { get; set; }
        public int partFastening { get; set; }
        
        public DataService(CFS_NX dialog)
        {
            _Dialog = dialog;            
        }

        public void Initialise_Data()
        {
            _Dialog.selectionGroup.Show = false;
            _Dialog.machineGroup.Show = false;
            _Dialog.opsOrGroupSelection.Show = false;
            
            _attrbutes = new Attrbutes(_Session);
            _uiCtrl = new SelectionCntrl();

            _Dialog.Rockwell.Value = 20.0;
            partHardness = _Dialog.Rockwell.Value;
            _Dialog.integer0.Value = 1;
            partFastening = _Dialog.integer0.Value;

            _Dialog.machine.SetEnumMembers(_attrbutes.ListMachine.ToArray());
            if (string.IsNullOrEmpty(_Dialog.machine.ValueAsString)) _Dialog.machine.ValueAsString = _attrbutes.ListMachine.FirstOrDefault();
            machineName = _Dialog.machine.ValueAsString;
            
            _Dialog.materialList.SetEnumMembers(_attrbutes.ListMaterial.ToArray());
            if (string.IsNullOrEmpty(_Dialog.materialList.ValueAsString)) _Dialog.materialList.ValueAsString = _attrbutes.ListMaterial.FirstOrDefault();
            partMaterial = _Dialog.materialList.ValueAsString;
        }

        public void updateCb(UIBlock block)
        {  
            if (block == _Dialog.machine)
            {
                machineName = string.Format(_Dialog.machine.ValueAsString, Enc);
            }
            else if (block == _Dialog.Rockwell)
            {                
                partHardness = _Dialog.Rockwell.Value;
            }
            else if (block == _Dialog.materialList)
            {
                partMaterial = string.Format(_Dialog.materialList.ValueAsString, Enc);
            }            
            else if (block == _Dialog.integer0)
            {
                partFastening = _Dialog.integer0.Value;
            }
            else if (block == _Dialog.applyFNS)
            {
                /*  For future release
                List<TaggedObject> TaggedObjects = new List<TaggedObject>();
                int selectedCount;
                Tag[] selectedTags;
                if (Ufs == null) return;
                Ufs.UiOnt.AskSelectedNodes(out selectedCount, out selectedTags);
                var taggedObjects = selectedTags.Select(NXObjectManager.Get).ToList();
                var objects = taggedObjects.Where(obj => obj is NCGroup || obj is NXOpen.CAM.Operation).ToList();

                if (selectedTags.Count() == 0) return;

                TaggedObjects = objects;
                //_uiCtrl.getSelectedandLoad(TaggedObjects, LibRecList);
                */
            }
            else if (block == _Dialog.saveFNS)
            {
                List<TaggedObject> TaggedObjects = new List<TaggedObject>();
                int selectedCount;
                Tag[] selectedTags;

                if (Ufs == null) return;
                
                    Ufs.UiOnt.AskSelectedNodes(out selectedCount, out selectedTags);
                    var taggedObjects = selectedTags.Select(NXObjectManager.Get).ToList();

                if (selectedTags.Count() == 0) return;

                var objects = taggedObjects.Where(obj => obj is NCGroup || obj is NXOpen.CAM.Operation).ToList();
                TaggedObjects = objects;
                _uiCtrl.getSelectedandSave(TaggedObjects, this);                
            }     
        }
    }
}
