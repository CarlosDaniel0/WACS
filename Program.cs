using WACS.Entities;

try
{
  AutoSheet.CleanSpreadsheets();
}
catch (Exception e)
{
  AutoSheet.Log($"Error - {e.Message}, {e.StackTrace}, {e.Source}");
}
await Console.Out.WriteLineAsync("Auto Clean Spreadsheets runned");
