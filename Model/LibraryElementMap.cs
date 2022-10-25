using CsvHelper.Configuration;

namespace CustomFeedsSpeedsLibrary.Model
{
    public class LibraryElementMap : ClassMap<LibraryElement> 
    {
        public LibraryElementMap()
        {
            Map(x => x.ToolType).Name("Тип Инструмента");
            Map(x => x.ToolSubType).Name("Подтип Инструмента");
            Map(x => x.toolName).Name("Имя Инструмента");
            Map(x => x.idInsert).Name("id Пластины");
            Map(x => x.idTool).Name("id инструмента");
            Map(x => x.holderName).Name("Держатель");
            Map(x => x.toolDiameter).Name("Диаметр Инструмента");
            Map(x => x.toolZCount).Name("Кол-во кромок");
            Map(x => x.toolOffset).Name("Вылет");
            Map(x => x.diamToOffset).Name("Диаметр/Вылет");
            Map(x => x.fZ).Name("fZ");
            Map(x => x.Vc).Name("Vc");
            Map(x => x.CutWidth).Name("Ширина резания");
            Map(x => x.CutStep).Name("Шаг");
            Map(x => x.operationType).Name("Тип операции");
            Map(x => x.globalMachineName).Name("Станок");
            Map(x => x.globalPartMat).Name("Материал");
            Map(x => x.globalPartHard).Name("Твердость HRc");
            Map(x => x.globalPartFastening).Name("Жесткость крепления");
            Map(x => x.userName).Name("Пользователь");
            Map(x => x.partName).Name("Номер детали");
            Map(x => x.opGroupName).Name("Опер/Группа");
            Map(x => x.DateTime).Name("Дата");
        }        
    }        
}

