using WACS.Entities;
using Quartz;

namespace WACS.Core
{
    class Job: IJob
    {
         public async Task Execute(IJobExecutionContext context)
         {
            try {
                AutoSheet.CleanSpreadsheets();
            } catch (Exception e) {
                AutoSheet.Log($"Error - {e.Message}");
            }
            await Console.Out.WriteLineAsync("Auto Clean Spreadsheets is running");
         }
    }
}