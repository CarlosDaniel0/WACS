using WACS.Entities;

try
{
  AutoSheet.CleanSpreadsheets();
}
catch (Exception e)
{
  AutoSheet.Log($"Error - {e.Message}");
}
await Console.Out.WriteLineAsync("Auto Clean Spreadsheets runned");
