using System.Diagnostics;

namespace WACS.Utils
{
    static class Tools
    {
        public static void CMD(string command)
        {
            Process process = new Process();
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                WindowStyle = ProcessWindowStyle.Hidden,
                FileName = "cmd.exe",
                Arguments = command
            };
            process.StartInfo = startInfo;
            process.Start();
        } 
    }
}