using System;
using System.IO;
using System.Globalization;
using WACS.Utils;
using System.Text.RegularExpressions;

namespace WACS.Entities
{
    class AutoSheet
    {
        public static void MoveFolderOfLastMonth()
        {
            var currentDate = DateTime.Now;
            var prevDate = currentDate.AddMonths(-2);
            string folder = GetFolder(prevDate);
            string source = String.Format(@"{0}{1}", Constraints.ROOT_SPREADSHEETS_PATH, folder);
            if (Directory.Exists(source)) 
            {
                string target = String.Format(@"{0}{1}", Constraints.STORED_SPREADSHEETS_PATH, folder.Split(" - ")[1]);
                CreateDirectory(target); 
                try  
                {  
                    Directory.Move(source, target + String.Format(@"\{0}", folder));  
                }  
                catch (IOException exp)  
                {  
                    Console.WriteLine(exp.Message);  
                }  
            }
        }

        public void CleanSpreadsheets()
        {
            if (VerifyExecutions())
            {                 
                string prevFolder = GetFolder(DateTime.Now.AddMonths(-1));
                string newfolder = GetFolder(DateTime.Now);
                string source = Constraints.ROOT_SPREADSHEETS_PATH + prevFolder;

                if (Directory.Exists(source)) 
                {
                    string target = Constraints.ROOT_SPREADSHEETS_PATH + newfolder;
                    CopyFilesRecursively(source, target);
                    Tools.CMD(String.Format(@"/C {0} ""{1}""", Constraints.CLI_PATH, target));
                    Log();
                    MoveFolderOfLastMonth(); 
                }
            }
        }

        public static bool VerifyExecutions()
        {
            DateTime today = DateTime.Now; 
            string log_path = String.Format("{0}{1}", Constraints.LOG_PATH, "execution.txt");
            if (File.Exists(log_path))
            {
                String last = File.ReadLines(log_path).Last();
                var date_rx = new Regex(@"((\d{2})/(\d{2})/(\d{4}) (\d{2}):(\d{2}):(\d{2}))");
                var date_match = date_rx.Match(last);
                if (date_match.Success)
                {
                    DateTime lastDate = DateTime.Parse(date_match.Groups[0].Value, new CultureInfo("pt-BR"));
                    TimeSpan diff = today - lastDate;
                    return diff.TotalDays > 4;
                }
                return false;
            }
            return true;
        }

        private static string GetFolder(DateTime date)
        {
            string month = date.ToString("MMMM", new CultureInfo("pt-BR"));
            month = month.ToUpper();
            return $"{month} - {date.Year}"; 
        }

        private static void CopyFilesRecursively(string sourcePath, string targetPath)
        {
            foreach (string dirPath in Directory.GetDirectories(sourcePath, "*", SearchOption.AllDirectories))
            {
                Directory.CreateDirectory(dirPath.Replace(sourcePath, targetPath));
            }

            foreach (string newPath in Directory.GetFiles(sourcePath, "*.*",SearchOption.AllDirectories))
            {
                File.Copy(newPath, newPath.Replace(sourcePath, targetPath), true);
            }
        }

        private static void CreateDirectory(string path)
        {
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
        }

        public static void Log(string status = "SUCCESS")
        {
            string today = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
            string message = $"{today} Running Service of Auto Clean Spreadsheets {status}";
            CreateDirectory(Constraints.LOG_PATH);
            using (StreamWriter sw = new StreamWriter(String.Format("{0}{1}", Constraints.LOG_PATH, "execution.txt"), true))
            {
                sw.WriteLine(message);
            }
            
        }
    }
}