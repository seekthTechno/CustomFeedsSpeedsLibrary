using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using NXOpen;

namespace CustomFeedsSpeedsLibrary.Data
{
    public class Attrbutes
    {
        public string AttrFilePath;

        private readonly Part _part1;
        //private NxSession _session;
        private List<string> _listMaterialList;
        private List<string> _listMachine;
        public string ROOT_PATH { get; set; }
        public string ROOT_PATH_TXT { get; set; }        
        
        public Attrbutes(Session session)
        {
            if (session == null) throw new Exception("Сессия NX не запущена!");
            ROOT_PATH = GetRootDirectory();

            var directories = Directory.GetDirectories(ROOT_PATH, "txt", SearchOption.AllDirectories);
            if (!directories.Any()) throw new Exception("Не найдена корневая директория с тектовыми файлами программы!");
            ROOT_PATH_TXT = directories.First();
            GetAttributeFromFiles();
            _part1 = session.Parts.Work;            
        }

        private static string GetRootDirectory()
        {
            var directory = AppDomain.CurrentDomain.BaseDirectory;
            var directoryInfo = new DirectoryInfo(directory).Parent;
            return directoryInfo != null ? directoryInfo.FullName : directory;
        }

        public List<string> ListMaterial
        {
            get { return _listMaterialList; }
            set { _listMaterialList = value; }
        }

        public List<string> ListMachine
        {
            get { return _listMachine; }
            set { _listMachine = value; }
        }
         
        private Part _Part
        {
            get { if (_part1 == null) throw new Exception("Не найдена открытая деталь!"); return _part1; }
        }
               
        public void GetAttributeFromFiles()
        {
            var filePath = ROOT_PATH_TXT;
            if (string.IsNullOrEmpty(filePath)) throw new Exception("Не удалось найти директорию с текстовыми файлами настройки!");

            ListMaterial = GetListFromFiles(Directory.GetFiles(filePath, "материал*.txt", SearchOption.AllDirectories));
            ListMachine = GetListFromFiles(Directory.GetFiles(filePath, "станок.txt", SearchOption.AllDirectories));
        }

        private static List<string> GetListFromFiles(string[] files)
        {
            return files.Any()
                ? files.Select(f => File.ReadAllLines(f, Encoding.Default).Where(s => !String.IsNullOrEmpty(s)).Select(s => s.Split(',').First())).SelectMany(l => l).ToList()
                : new List<string>();
        }
                
        private static List<string> GetAttributeFilter(string filename)
        {
            var ret = new List<string>();
            if (File.Exists(filename)) ret.AddRange(File.ReadAllLines(filename));
            return ret;
        }
    }
}
