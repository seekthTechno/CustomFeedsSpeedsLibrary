using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CsvHelper;
using CsvHelper.Configuration;
using CustomFeedsSpeedsLibrary.Model;
using NXOpen;
using System.IO;
using System.Globalization;

namespace CustomFeedsSpeedsLibrary.Services
{
    public class LibraryRecordUpload
    {       
        private static UI _UI = UI.GetUI();
        private static Session ses = Session.GetSession();
                
        private string filePath = Properties.Settings.Default.BasePath + Properties.Settings.Default.BaseFileName;
        private string fileDir = Properties.Settings.Default.BasePath;

        public void save2file(List<LibraryElement> LRLsaveList)
        {                            
            ///// XML
            
            ///// CSV            
            if (!Directory.Exists(fileDir))
            {
                Directory.CreateDirectory(fileDir);
            }            
                
                if (File.Exists(filePath))
                {

                var config = new CsvConfiguration(CultureInfo.InvariantCulture)
                {                    
                    HasHeaderRecord = false                        
                };

                List<LibraryElement> fromfile = new List<LibraryElement>();   
                try
                {
                    using (var stream = File.Open(filePath, FileMode.Open))
                    using (var reader = new StreamReader(stream, Encoding.Default))
                    using (var data = new CsvReader(reader, CultureInfo.InvariantCulture))
                    { 
                        data.Read();
                        while (data.Read())
                        {
                            fromfile.Add(new LibraryElement() {
                                ToolType = data.GetField(0),
                                ToolSubType = data.GetField(1),
                                toolName = string.Format(data.GetField(2), Encoding.Default),
                                idInsert = data.GetField(3),
                                idTool = data.GetField(4),
                                holderName = string.Format(data.GetField(5), Encoding.Default),
                                toolDiameter = data.GetField(6),
                                toolZCount = data.GetField(7),
                                toolOffset = data.GetField(8),
                                diamToOffset = data.GetField(9),
                                fZ = data.GetField(10),
                                Vc = data.GetField(11),
                                CutWidth = data.GetField(12),
                                CutStep = data.GetField(13),
                                operationType = data.GetField(14),
                                globalMachineName = string.Format(data.GetField(15), Encoding.Default),
                                globalPartMat = string.Format(data.GetField(16), Encoding.Default),
                                globalPartHard = data.GetField(17),
                                globalPartFastening = data.GetField(18)
                            });
                        }
                        data.Dispose();
                        
                        foreach (LibraryElement a in LRLsaveList.ToList())
                        {
                            var c = fromfile.Any(x => x.CutStep == a.CutStep &&
                            x.CutWidth == a.CutWidth &&
                            x.diamToOffset == a.diamToOffset &&
                            x.fZ == a.fZ &&
                            x.holderName == a.holderName &&
                            x.idInsert == a.idInsert &&
                            x.idTool == a.idTool &&
                            x.globalMachineName == a.globalMachineName &&
                            x.globalPartFastening == a.globalPartFastening &&
                            x.globalPartHard == a.globalPartHard &&
                            x.globalPartMat == a.globalPartMat &&
                            x.operationType == a.operationType &&
                            x.toolDiameter == a.toolDiameter &&
                            x.toolName == a.toolName &&
                            x.toolOffset == a.toolOffset &&
                            x.ToolSubType == a.ToolSubType &&
                            x.ToolType== a.ToolType &&
                            x.toolZCount == a.toolZCount && 
                            x.Vc == a.Vc);

                            if (c)
                            {
                                ses.ListingWindow.Open();
                                ses.ListingWindow.WriteLine("Запись для " + a.toolName.ToUpper() + " уже существует");
                                LRLsaveList.Remove(a);  
                            }
                           
                        }
                        
                    }
                }
                catch (Exception ex) {
                    int answer = _UI.NXMessageBox.Show("Файл заблокирован", NXMessageBox.DialogType.Error, ex.ToString());
                }

                if (LRLsaveList.Count == 0) return;

                try
                {
                    using (var stream = File.Open(filePath, FileMode.Append))
                    using (var writer = new StreamWriter(stream, Encoding.Default))
                    using (var csv = new CsvWriter(writer, config))
                    {
                        csv.Context.RegisterClassMap<LibraryElementMap>();
                        foreach (var a in LRLsaveList)
                        {
                            csv.NextRecord();
                            csv.WriteRecord(a);
                        }
                        writer.Flush();
                    }
                }
                catch (Exception ex)
                {
                    int answer = _UI.NXMessageBox.Show("Файл заблокирован", NXMessageBox.DialogType.Error, "Файл базы скорей всего редактируется другим пользователем");
                }

                }
                else
                {
                    var config = new CsvConfiguration(CultureInfo.InvariantCulture)
                    { };
                    try
                    {
                    using (var stream = File.Open(filePath, FileMode.CreateNew))
                    using (var writer = new StreamWriter(stream, Encoding.Default))
                        using (var csv = new CsvWriter(writer, config))
                        {
                        csv.Context.RegisterClassMap<LibraryElementMap>();
                        csv.WriteHeader<LibraryElement>();
                                                                           
                            foreach (var a in LRLsaveList)
                            {
                            csv.NextRecord();
                            csv.WriteRecord(a);
                            }
                            writer.Flush();
                        }
                    }
                    catch (Exception ex)
                    {
                        int answer = _UI.NXMessageBox.Show("Файл заблокирован", NXMessageBox.DialogType.Error, ex.ToString());
                    }

                }
           
        }
    }
          
}
