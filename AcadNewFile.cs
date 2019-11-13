using EPDM.Interop.epdm;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Console_Test
{
    class AcadNewFile
    {
        static IEdmVault7 vault;
        static bool NeedToDoIt = false;
        static IEdmFolder5 lastfolder = null;

        internal static void CreateNewFile()
        {
            Stopwatch sw = new Stopwatch();            

            sw.Start();

            CreateDWG();

            sw.Stop();

            TimeSpan ts2 = sw.Elapsed;
            Console.WriteLine($"Затраченное время {ts2.TotalMilliseconds} мс");
        }

        internal static void CreateDWG()
        {   // Этот метод создает новый файл копируя шаблон
            Console.WriteLine("Создаем новый dwg");

            vault = new EdmVault5();
            if (!vault.IsLoggedIn) { vault.LoginAuto("PDM_GORELTEX", 0); }

            string dir1 = @"TEST";
            string dir2 = @"Чертежи ПО ИСХ.НОМЕРАМ";
            string path = Path.Combine(vault.RootFolderPath, dir2);

            string TemplatePath = Path.Combine(vault.RootFolderPath, "Шаблон.dwg");
            IEdmFile5 Template = vault.GetFileFromPath(TemplatePath, out IEdmFolder5 folder);

            string NewName = "";
            string LastFileName = GetLastFileName(path);

            if (int.TryParse(LastFileName, out int numberOfDwg ))
            {
                numberOfDwg++;
                if (NeedToDoIt)
                {
                    string newFolderName = String.Format("{0}.{1}", DateTime.Now.Month, numberOfDwg);
                    lastfolder = lastfolder.ParentFolder.AddFolder(0, newFolderName);
                    NeedToDoIt = false;
                }
                if (numberOfDwg.ToString().Length == 5)
                    NewName = String.Format($"{numberOfDwg:00000}.{DateTime.Now:MM.yy}.01.dwg");
                else NewName = String.Format($"{numberOfDwg:0000}.{DateTime.Now:MM.yy}.01.dwg");
                Console.WriteLine($"Имя нового файла {NewName}");
            }

            //try { lastfolder.CopyFile(Template.ID, folder.ID, 0, NewName); }
            //catch 
            //{
            //    numberOfDwg++;
            //    if (numberOfDwg.ToString().Length == 5)
            //        NewName = String.Format($"{numberOfDwg:00000}.{DateTime.Now:MM.yy}.01.dwg");
            //    else NewName = String.Format($"{numberOfDwg:0000}.{DateTime.Now:MM.yy}.01.dwg");
            //    Console.WriteLine($"Имя нового файла {NewName}");
            //}
            
        }

        internal static string GetLastFileName(string path)
        {   // Этот метод ищет имя последнего файла в папке
            Console.WriteLine("Ищем имя последнего файла");

            lastfolder = GetLastFolder(path);

            IEdmSearch6 Search = (IEdmSearch6)vault.CreateSearch();
            Search.FindFolders = false;
            Search.FindFiles = true;
            Search.StartFolderID = lastfolder.ID;

            SearchTestPDM.GetSearchResults(Search, out List<IEdmSearchResult5> Results);

            var SortedResults = from result in Results
                                orderby result.Name descending
                                select result;

            IEdmSearchResult5 lastresult = SortedResults.ToArray()[0] as IEdmSearchResult5;

            string lastName = lastresult.Name.Substring(0, lastresult.Name.Length - 13);

            Console.WriteLine( $"Имя последнего файла: {lastresult.Name}");

            return lastName;
        }

        private static IEdmFolder5 GetLastFolder(string path)
        {   // Этот метод ищет по указанному пути папку с текущим годом
            // если папки с годом нет, создает ее с именем текущего года
            // а в ней создает папку "01.0001" для чертежей.
            // если папка с годом существует и в ней существует подпапка с месяцем\
            // идет проверка соответствует ли последняя подпапка текущему месяцу.
            // если подпапка соотвествует, флаг о необходимости создания новой папки становится true
            // методу возвращается значение последней подпапки в папке с годом.
            
            Console.WriteLine("Ищем имя последней папки");

            IEdmFolder5 lastfolder = null;

            string Year = DateTime.Now.Year.ToString();
            string Month = DateTime.Now.Month.ToString();
            string YearPath = Path.Combine(path, Year);

            IEdmFolder5 PathFolder = vault.GetFolderFromPath(path);
            IEdmFolder5 YearFolder = vault.GetFolderFromPath(YearPath);

            if (YearFolder == null)
            {
                PathFolder.AddFolder(0, Year);
                YearFolder = vault.GetFolderFromPath(YearPath);
                YearFolder.AddFolder(0, "01.0001");
                lastfolder = vault.GetFolderFromPath(Path.Combine(YearPath, "01.0001"));
            }
            else
            {
                IEdmPos5 pos = YearFolder.GetFirstSubFolderPosition();
                while (!pos.IsNull)
                {
                    lastfolder = YearFolder.GetNextSubFolder(pos);                    
                }
                if (lastfolder.Name.Substring(0, 2) != Month) NeedToDoIt = true;                
            }

            Console.WriteLine($"Имя последней папки {lastfolder.Name}");

            return lastfolder;
        }

        internal static void GetDirectory(string path)
        {   // Этот метод получает последнюю папку по указанному пути
            DirectoryInfo dir = new DirectoryInfo(path);
            DirectoryInfo[] folders = dir.GetDirectories();
            DirectoryInfo lastfolder = folders[folders.GetUpperBound(0)];
            string currentPath = lastfolder.FullName;
            Console.WriteLine(currentPath);
        }
    }
}
